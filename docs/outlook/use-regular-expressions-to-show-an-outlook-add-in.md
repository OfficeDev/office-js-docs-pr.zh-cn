---
title: 使用正则表达式激活规则显示加载项
description: 了解如何为 Outlook 上下文加载项使用正则表达式激活规则。
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: d334ba6b2e0f044fc8d876cd6edd218743ccb390
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348851"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a><span data-ttu-id="9ebcd-103">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="9ebcd-103">Use regular expression activation rules to show an Outlook add-in</span></span>

<span data-ttu-id="9ebcd-104">可以将正则表达式规则指定为在邮件的特定字段中找到匹配项时激活[上下文外接程序](contextual-outlook-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-104">You can specify regular expression rules to have a [contextual add-in](contextual-outlook-add-ins.md) activated when a match is found in specific fields of the message.</span></span> <span data-ttu-id="9ebcd-105">上下文外接程序仅在阅读模式下激活，Outlook 不会在用户撰写某个项目时激活上下文外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-105">Contextual add-ins activate only in read mode, Outlook does not activate contextual add-ins when the user is composing an item.</span></span> <span data-ttu-id="9ebcd-106">还有其他一些情况Outlook激活外接程序，例如，数字签名项目。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-106">There are also other scenarios where Outlook does not activate add-ins, for example, digitally signed items.</span></span> <span data-ttu-id="9ebcd-107">有关详细信息，请参阅 [Outlook 外接程序的激活规则](activation-rules.md)。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-107">For more information, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

<span data-ttu-id="9ebcd-108">你可以将正则表达式指定为外接程序 XML 清单中的 [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) 规则或 [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) 规则的一部分。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-108">You can specify a regular expression as part of an [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule or [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule in the add-in XML manifest.</span></span> <span data-ttu-id="9ebcd-109">在 [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) 扩展点中指定了这些规则。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-109">The rules are specified in a [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) extension point.</span></span>

<span data-ttu-id="9ebcd-110">Outlook 基于客户端计算机上浏览器所使用的 JavaScript 解释器的规则计算正则表达式。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-110">Outlook evaluates regular expressions based on the rules for the JavaScript interpreter used by the browser on the client computer.</span></span> <span data-ttu-id="9ebcd-111">Outlook 支持所有 XML 处理器也支持的相同特殊字符列表。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-111">Outlook supports the same list of special characters that all XML processors also support.</span></span> <span data-ttu-id="9ebcd-112">下表列出了这些特殊字符。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-112">The following table lists these special characters.</span></span> <span data-ttu-id="9ebcd-113">你可以通过为相应字符指定转义序列以在正则表达式中使用这些字符，如下表中所述。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-113">You can use these characters in a regular expression by specifying the escaped sequence for the corresponding character, as described in the following table.</span></span>

<br/>

|<span data-ttu-id="9ebcd-114">字符</span><span class="sxs-lookup"><span data-stu-id="9ebcd-114">Character</span></span>|<span data-ttu-id="9ebcd-115">说明</span><span class="sxs-lookup"><span data-stu-id="9ebcd-115">Description</span></span>|<span data-ttu-id="9ebcd-116">要使用的转义序列</span><span class="sxs-lookup"><span data-stu-id="9ebcd-116">Escape sequence to use</span></span>|
|:-----|:-----|:-----|
|`"`|<span data-ttu-id="9ebcd-117">双引号</span><span class="sxs-lookup"><span data-stu-id="9ebcd-117">Double quotation mark</span></span>|`&quot;`|
|`&`|<span data-ttu-id="9ebcd-118">与号</span><span class="sxs-lookup"><span data-stu-id="9ebcd-118">Ampersand</span></span>|`&amp;`|
|`'`|<span data-ttu-id="9ebcd-119">撇号</span><span class="sxs-lookup"><span data-stu-id="9ebcd-119">Apostrophe</span></span>|`&apos;`|
|`<`|<span data-ttu-id="9ebcd-120">小于号</span><span class="sxs-lookup"><span data-stu-id="9ebcd-120">Less-than sign</span></span>|`&lt;`|
|`>`|<span data-ttu-id="9ebcd-121">大于号</span><span class="sxs-lookup"><span data-stu-id="9ebcd-121">Greater-than sign</span></span>|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="9ebcd-122">ItemHasRegularExpressionMatch 规则</span><span class="sxs-lookup"><span data-stu-id="9ebcd-122">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="9ebcd-123">`ItemHasRegularExpressionMatch` 规则对于基于受支持属性的特定值控制外接程序的激活很有用。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-123">An  `ItemHasRegularExpressionMatch` rule is useful in controlling activation of an add-in based on specific values of a supported property.</span></span> <span data-ttu-id="9ebcd-124">`ItemHasRegularExpressionMatch` 规则具有以下属性。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-124">The `ItemHasRegularExpressionMatch` rule has the following attributes.</span></span>

<br/>

|<span data-ttu-id="9ebcd-125">属性名</span><span class="sxs-lookup"><span data-stu-id="9ebcd-125">Attribute name</span></span>|<span data-ttu-id="9ebcd-126">说明</span><span class="sxs-lookup"><span data-stu-id="9ebcd-126">Description</span></span>|
|:-----|:-----|
|`RegExName`|<span data-ttu-id="9ebcd-127">指定正则表达式的名称，以便能够在外接程序的代码中引用该表达式。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-127">Specifies the name of the regular expression so that you can refer to the expression in the code for your add-in.</span></span>|
|`RegExValue`|<span data-ttu-id="9ebcd-128">指定将对其求值的正则表达式，以确定是否应显示外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-128">Specifies the regular expression that will be evaluated to determine whether the add-in should be shown.</span></span>|
|`PropertyName`|<span data-ttu-id="9ebcd-129">指定正则表达式进行计算所依据的属性名称。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-129">Specifies the name of the property that the regular expression will be evaluated against.</span></span> <span data-ttu-id="9ebcd-130">允许的值为 `BodyAsHTML`、`BodyAsPlaintext`、`SenderSMTPAddress` 和 `Subject`。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-130">The allowed values are `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress`, and `Subject`.</span></span><br/><br/><span data-ttu-id="9ebcd-131">如果指定 `BodyAsHTML`，则 Outlook 只会在项目正文为 HTML 时应用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-131">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="9ebcd-132">否则，Outlook 将不会返回该正则表达式的匹配项。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-132">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="9ebcd-133">如果指定 `BodyAsPlaintext`，则 Outlook 将始终对项目正文应用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-133">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="9ebcd-134">**注释：** 如果指定 `Rule` 元素的 `Highlight` 属性，则必须将 `PropertyName` 属性设为 `BodyAsPlaintext`。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-134">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span>|
|`IgnoreCase`|<span data-ttu-id="9ebcd-135">指定当匹配由 `RegExName` 指定的正则表达式时是否忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-135">Specifies whether to ignore case when matching the regular expression specified by `RegExName`.</span></span>|
| `Highlight` | <span data-ttu-id="9ebcd-136">指定客户端应如何突出显示匹配的文本。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-136">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="9ebcd-137">此元素仅适用于 `ExtensionPoint` 元素中的 `Rule` 元素。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-137">This element can only be applied to `Rule` elements within `ExtensionPoint` elements.</span></span> <span data-ttu-id="9ebcd-138">可以是以下值之一：`all` 或 `none`。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-138">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="9ebcd-139">如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-139">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="9ebcd-140">**注释：** 如果指定 `Rule` 元素的 `Highlight` 属性，则必须将 `PropertyName` 属性设为 `BodyAsPlaintext`。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-140">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span> |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a><span data-ttu-id="9ebcd-141">在规则中使用正则表达式的最佳做法</span><span class="sxs-lookup"><span data-stu-id="9ebcd-141">Best practices for using regular expressions in rules</span></span>

<span data-ttu-id="9ebcd-142">在使用正则表达式时，请特别注意以下几点。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-142">Pay special attention to the following when you use regular expressions.</span></span>

- <span data-ttu-id="9ebcd-143">如果在项目的正文中指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-143">If you specify an `ItemHasRegularExpressionMatch` rule on the body of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item.</span></span> <span data-ttu-id="9ebcd-144">使用正则表达式（如 `.*`）来尝试获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-144">Using a regular expression such as `.*` to attempt to obtain the entire body of an item does not always return the expected results.</span></span>
- <span data-ttu-id="9ebcd-145">一个浏览器上返回的纯文本正文与另一个浏览器上返回的纯文本正文可能略有不同。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-145">The plain text body returned on one browser can be different in subtle ways on another.</span></span> <span data-ttu-id="9ebcd-146">如果使用含有 `BodyAsPlaintext` 的 `ItemHasRegularExpressionMatch` 规则作为 `PropertyName` 属性，请在你的外接程序支持的所有浏览器上测试正则表达式。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-146">If you use an `ItemHasRegularExpressionMatch` rule with `BodyAsPlaintext` as the `PropertyName` attribute, test your regular expression on all the browsers that your add-in supports.</span></span>

    <span data-ttu-id="9ebcd-147">因为不同的浏览器获取所选项目的文本正文的方法不同，所以应确保你的正则表达式支持正文文本部分所返回的细微差异。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-147">Because different browsers use different ways to obtain the text body of a selected item, you should make sure that your regular expression supports the subtle differences that can be returned as part of the body text.</span></span> <span data-ttu-id="9ebcd-148">例如，一些浏览器（如 Internet Explorer 9）使用 DOM 的 `innerText` 属性，而其他浏览器（如 Firefox）使用.`.textContent()` 方法来获取项目的文本正文。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-148">For example, some browsers such as Internet Explorer 9 uses the `innerText` property of the DOM, and others such as Firefox uses the `.textContent()` method to obtain the text body of an item.</span></span> <span data-ttu-id="9ebcd-149">同样，不同浏览器所返回的换行符也可能不同：在 Internet Explorer 上返回的换行符为 `\r\n`，而在 Firefox 和 Chrome 上返回的换行符为 `\n`。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-149">Also, different browsers may return line breaks differently: a line break is `\r\n` on Internet Explorer, and `\n` on Firefox and Chrome.</span></span> <span data-ttu-id="9ebcd-150">有关详细信息，请参阅 [W3C DOM 兼容性 - HTML](https://quirksmode.org/dom/html/)。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-150">For more information, se [W3C DOM Compatibility - HTML](https://quirksmode.org/dom/html/).</span></span>

- <span data-ttu-id="9ebcd-151">Outlook 富客户端与 Outlook 网页版或 Outlook Mobile 之间的项目的 HTML 正文略有不同。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-151">The HTML body of an item is slightly different between an Outlook rich client, and Outlook on the web or Outlook mobile.</span></span> <span data-ttu-id="9ebcd-152">请仔细定义正则表达式。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-152">Define your regular expressions carefully.</span></span>

- <span data-ttu-id="9ebcd-153">根据 Outlook 客户端、设备类型或要应用正则表达式的属性，在设计正则表达式作为激活规则时，您应该了解每个客户端的其他最佳实践和限制。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-153">Depending on the Outlook client, type of device, or property that a regular expression is being applied on, there are other best practices and limits for each of the clients that you should be aware of when designing regular expressions as activation rules.</span></span> <span data-ttu-id="9ebcd-154">有关详细信息，请参阅 [Outlook 外接程序的激活和 JavaScript API 的限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-154">See [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) for details.</span></span>

### <a name="examples"></a><span data-ttu-id="9ebcd-155">示例</span><span class="sxs-lookup"><span data-stu-id="9ebcd-155">Examples</span></span>

<span data-ttu-id="9ebcd-156">以下 `ItemHasRegularExpressionMatch` 规则将在发件人的 SMTP 电子邮件地址与 `@contoso` 匹配（不管是大写还是小写字符）时激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-156">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever the sender's SMTP email address matches `@contoso`, regardless of uppercase or lowercase characters.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

<span data-ttu-id="9ebcd-157">以下是使用 `IgnoreCase` 属性指定同一正则表达式的另一种方式。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-157">The following is another way to specify the same regular expression using the  `IgnoreCase` attribute.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

<span data-ttu-id="9ebcd-158">以下 `ItemHasRegularExpressionMatch` 规则将在股票代号包含在当前项目的正文中时激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-158">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever a stock symbol is included in the body of the current item.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="9ebcd-159">ItemHasKnownEntity 规则</span><span class="sxs-lookup"><span data-stu-id="9ebcd-159">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="9ebcd-160">`ItemHasKnownEntity` 规则根据所选项目的主题或正文中是否存在实体来激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-160">An `ItemHasKnownEntity` rule activates an add-in based on the existence of an entity in the subject or body of the selected item.</span></span> <span data-ttu-id="9ebcd-161">[EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) 类型定义受支持的实体。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-161">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) type defines the supported entities.</span></span> <span data-ttu-id="9ebcd-162">在 `ItemHasKnownEntity` 规则中应用正则表达式，可为基于实体（例如，一组特定的 URL，或含有某个区号的电话号码）的值子集进行的激活提供便利。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-162">Applying a regular expression on an `ItemHasKnownEntity` rule provides the convenience where activation is based on a subset of values for an entity (for example, a specific set of URLs, or telephone numbers with a certain area code).</span></span>

> [!NOTE]
> <span data-ttu-id="9ebcd-163">Outlook 只能提取用英语编写的实体字符串，无论清单中指定的默认区域设置如何。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-163">Outlook can only extract entity strings in English regardless of the default locale specified in the manifest.</span></span> <span data-ttu-id="9ebcd-164">仅邮件支持 `MeetingSuggestion` 实体类型；约会不支持该类型。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-164">Only messages support the `MeetingSuggestion` entity type; appointments do not.</span></span> <span data-ttu-id="9ebcd-165">你无法从“已发送邮件”文件夹的邮件中提取实体，也不能使用 `ItemHasKnownEntity` 规则来激活“已发送邮件”文件夹中邮件的外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-165">You cannot extract entities from items in the **Sent Items** folder, nor can you use an `ItemHasKnownEntity` rule to activate an add-in for items in the **Sent Items** folder.</span></span>

<span data-ttu-id="9ebcd-166">`ItemHasKnownEntity` 规则支持下表中的属性。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-166">The `ItemHasKnownEntity` rule supports the attributes in the following table.</span></span> <span data-ttu-id="9ebcd-167">请注意，尽管在 `ItemHasKnownEntity` 规则中指定正则表达式是可选项，如果选择使用正则表达式作为实体筛选器，则必须同时指定 `RegExFilter` 和 `FilterName` 属性。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-167">Note that while specifying a regular expression is optional in an `ItemHasKnownEntity` rule, if you choose to use a regular expression as an entity filter, you must specify both the `RegExFilter` and `FilterName` attributes.</span></span>

<br/>

|<span data-ttu-id="9ebcd-168">属性名</span><span class="sxs-lookup"><span data-stu-id="9ebcd-168">Attribute name</span></span>|<span data-ttu-id="9ebcd-169">说明</span><span class="sxs-lookup"><span data-stu-id="9ebcd-169">Description</span></span>|
|:-----|:-----|
|`EntityType`|<span data-ttu-id="9ebcd-170">指定若想规则计算结果为 `true` 而必须存在的实体类型。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-170">Specifies the type of entity that must be found for the rule to evaluate to `true`.</span></span> <span data-ttu-id="9ebcd-171">请使用多个规则来指定多个类型的实体。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-171">Use multiple rules to specify multiple types of entities.</span></span>|
|`RegExFilter`|<span data-ttu-id="9ebcd-172">指定用于进一步筛选由 `EntityType` 指定的实体实例的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-172">Specifies a regular expression that further filters instances of the entity specified by `EntityType`.</span></span>|
|`FilterName`|<span data-ttu-id="9ebcd-173">指定由 `RegExFilter` 指定的正则表达式的名称，以便稍后可通过代码引用它。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-173">Specifies the name of the regular expression specified by `RegExFilter`, so that it is subsequently possible to refer to it by code.</span></span>|
|`IgnoreCase`|<span data-ttu-id="9ebcd-174">指定当匹配由 `RegExFilter` 指定的正则表达式时是否忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-174">Specifies whether to ignore case when matching the regular expression specified by `RegExFilter`.</span></span>|

### <a name="examples"></a><span data-ttu-id="9ebcd-175">示例</span><span class="sxs-lookup"><span data-stu-id="9ebcd-175">Examples</span></span>

<span data-ttu-id="9ebcd-176">下面的 `ItemHasKnownEntity` 规则将在当前项目的主题或正文中存在 URL 且该 URL 包含字符串 `youtube` 时激活外接程序，而不考虑字符串的大小写。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-176">The following `ItemHasKnownEntity` rule activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string `youtube`, regardless of the case of the string.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a><span data-ttu-id="9ebcd-177">在代码中使用正则表达式结果</span><span class="sxs-lookup"><span data-stu-id="9ebcd-177">Using regular expression results in code</span></span>

<span data-ttu-id="9ebcd-178">可以通过对当前项使用下列方法获取正则表达式的匹配项。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-178">You can obtain matches to a regular expression by using the following methods on the current item.</span></span>

- <span data-ttu-id="9ebcd-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 为在外接程序的 `ItemHasRegularExpressionMatch` 和 `ItemHasKnownEntity` 规则中指定的所有正则表达式返回当前项目中的匹配项。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for all regular expressions specified in `ItemHasRegularExpressionMatch` and `ItemHasKnownEntity` rules of the add-in.</span></span>

- <span data-ttu-id="9ebcd-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 为外接程序的 `ItemHasRegularExpressionMatch` 规则中指定的已标识正则表达式返回当前项目中的匹配项。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for the identified regular expression specified in an `ItemHasRegularExpressionMatch` rule of the add-in.</span></span>

- <span data-ttu-id="9ebcd-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 对于包含在外接程序的 `ItemHasKnownEntity` 规则中指定的已标识正则表达式匹配项的实体，将返回完整实例。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns entire instances of entities that contain matches for the identified regular expression specified in an `ItemHasKnownEntity` rule of the add-in.</span></span>

<span data-ttu-id="9ebcd-182">计算正则表达式时，匹配项将以数组对象的形式返回到你的外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-182">When the regular expressions are evaluated, the matches are returned to your add-in in an array object.</span></span> <span data-ttu-id="9ebcd-183">对于 `getRegExMatches`，该对象具有正则表达式名称的标识符。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-183">For `getRegExMatches`, that object has the identifier of the name of the regular expression.</span></span>

> [!NOTE]
> <span data-ttu-id="9ebcd-184">Outlook 不会在数组中以任何特定顺序返回匹配项。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-184">Outlook does not return matches in any particular order in the array.</span></span> <span data-ttu-id="9ebcd-185">另外，即使在同一邮箱中的同一项目上的每个客户端运行相同的外接程序，也不应假定匹配项返回的顺序与数组中返回的顺序相同。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-185">Also, you should not assume that matches are returned in the same order in this array even when you run the same add-in on each of these clients on the same item in the same mailbox.</span></span>

### <a name="examples"></a><span data-ttu-id="9ebcd-186">示例</span><span class="sxs-lookup"><span data-stu-id="9ebcd-186">Examples</span></span>

<span data-ttu-id="9ebcd-187">以下是包含 `ItemHasRegularExpressionMatch` 规则且具有名为 `videoURL` 的正则表达式的规则集合示例。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-187">The following is an example of a rule collection that contains an  `ItemHasRegularExpressionMatch` rule with a regular expression named `videoURL`.</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

<span data-ttu-id="9ebcd-188">以下示例使用当前项目的 `getRegExMatches` 将变量 `videos` 设置为上一个 `ItemHasRegularExpressionMatch` 规则的结果。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-188">The following example uses `getRegExMatches` of the current item to set a variable `videos` to the results of the preceding `ItemHasRegularExpressionMatch` rule.</span></span>

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

<span data-ttu-id="9ebcd-p119">多个匹配项将作为数组元素存储在该对象中。以下代码示例说明如何对名为  `reg1` 的正则表达式循环访问匹配项以生成将显示为 HTML 的字符串。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-p119">Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.</span></span>

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

<span data-ttu-id="9ebcd-191">以下是指定 `MeetingSuggestion` 实体和名为 `CampSuggestion` 的正则表达式的 `ItemHasKnownEntity` 规则的示例。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-191">The following is an example of an `ItemHasKnownEntity` rule that specifies the `MeetingSuggestion` entity and a regular expression named `CampSuggestion`.</span></span> <span data-ttu-id="9ebcd-192">Outlook 在检测到当前所选项目包含会议建议，并且主题或正文包含术语 `WonderCamp` 时将激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-192">Outlook activates the add-in if it detects that the currently selected item contains a meeting suggestion, and the subject or body contains the term `WonderCamp`.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

<span data-ttu-id="9ebcd-193">以下代码示例使用当前项目中的 `getFilteredEntitiesByName` 设置变量 `suggestions`，以获取针对上一个 `ItemHasKnownEntity` 规则检测到的一组会议建议。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-193">The following code example uses `getFilteredEntitiesByName` on the current item to set a variable `suggestions` to an array of detected meeting suggestions for the preceding `ItemHasKnownEntity` rule.</span></span>

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a><span data-ttu-id="9ebcd-194">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9ebcd-194">See also</span></span>

- <span data-ttu-id="9ebcd-195">[Outlook 外接程序：Contoso 订单编号](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - 基于正则表达式匹配项激活的示例上下文外接程序。</span><span class="sxs-lookup"><span data-stu-id="9ebcd-195">[Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - A sample contextual add-in that activates based on a regular expression match.</span></span>
- [<span data-ttu-id="9ebcd-196">创建适用于阅读窗体的 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="9ebcd-196">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="9ebcd-197">Outlook 外接程序的激活规则</span><span class="sxs-lookup"><span data-stu-id="9ebcd-197">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="9ebcd-198">Outlook 外接程序的激活和 JavaScript API 限制</span><span class="sxs-lookup"><span data-stu-id="9ebcd-198">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="9ebcd-199">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="9ebcd-199">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="9ebcd-200">.NET Framework 中的正则表达式的最佳做法</span><span class="sxs-lookup"><span data-stu-id="9ebcd-200">Best Practices for Regular Expressions in the .NET Framework</span></span>](/dotnet/standard/base-types/best-practices)
