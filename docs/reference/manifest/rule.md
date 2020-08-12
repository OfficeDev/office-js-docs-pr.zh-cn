---
title: 清单文件中的 Rule 元素
description: Rule 元素指定应为此上下文邮件外接程序计算的激活规则。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 79b97f2e442e9d8ce59d17467161b5b9b7a7252d
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641429"
---
# <a name="rule-element"></a><span data-ttu-id="f627e-103">Rule 元素</span><span class="sxs-lookup"><span data-stu-id="f627e-103">Rule element</span></span>

<span data-ttu-id="f627e-104">指定应针对此上下文邮件外接程序计算的激活规则。</span><span class="sxs-lookup"><span data-stu-id="f627e-104">Specifies the activation rules that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="f627e-105">**外接类型：** 邮件 (上下文) </span><span class="sxs-lookup"><span data-stu-id="f627e-105">**Add-in type:** Mail (contextual)</span></span>

## <a name="contained-in"></a><span data-ttu-id="f627e-106">包含于</span><span class="sxs-lookup"><span data-stu-id="f627e-106">Contained in</span></span>

- [<span data-ttu-id="f627e-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f627e-107">OfficeApp</span></span>](officeapp.md)
- <span data-ttu-id="f627e-108">[ExtensionPoint](extensionpoint.md) ([ **CustomPane** (弃用) ](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)， [**DetectedEntity**](extensionpoint.md#detectedentity)) </span><span class="sxs-lookup"><span data-stu-id="f627e-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span></span>

## <a name="attributes"></a><span data-ttu-id="f627e-109">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-109">Attributes</span></span>

| <span data-ttu-id="f627e-110">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-110">Attribute</span></span> | <span data-ttu-id="f627e-111">必需</span><span class="sxs-lookup"><span data-stu-id="f627e-111">Required</span></span> | <span data-ttu-id="f627e-112">说明</span><span class="sxs-lookup"><span data-stu-id="f627e-112">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="f627e-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="f627e-113">**xsi:type**</span></span> | <span data-ttu-id="f627e-114">是</span><span class="sxs-lookup"><span data-stu-id="f627e-114">Yes</span></span> | <span data-ttu-id="f627e-115">正在定义的规则类型。</span><span class="sxs-lookup"><span data-stu-id="f627e-115">The type of rule being defined.</span></span> |

<span data-ttu-id="f627e-116">此规则类型可以是下列类型之一。</span><span class="sxs-lookup"><span data-stu-id="f627e-116">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="f627e-117">ItemIs</span><span class="sxs-lookup"><span data-stu-id="f627e-117">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="f627e-118">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="f627e-118">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="f627e-119">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="f627e-119">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="f627e-120">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="f627e-120">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="f627e-121">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="f627e-121">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="f627e-122">ItemIs 规则</span><span class="sxs-lookup"><span data-stu-id="f627e-122">ItemIs rule</span></span>

<span data-ttu-id="f627e-123">定义一个在选定项为指定类型时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="f627e-123">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="f627e-124">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-124">Attributes</span></span>

| <span data-ttu-id="f627e-125">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-125">Attribute</span></span> | <span data-ttu-id="f627e-126">必需</span><span class="sxs-lookup"><span data-stu-id="f627e-126">Required</span></span> | <span data-ttu-id="f627e-127">说明</span><span class="sxs-lookup"><span data-stu-id="f627e-127">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="f627e-128">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="f627e-128">**ItemType**</span></span> | <span data-ttu-id="f627e-129">是</span><span class="sxs-lookup"><span data-stu-id="f627e-129">Yes</span></span> | <span data-ttu-id="f627e-p101">指定要匹配的项目类型。可以是 `Message` 或 `Appointment`。`Message` 项目类型包括电子邮件、会议请求、会议响应和会议取消。</span><span class="sxs-lookup"><span data-stu-id="f627e-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="f627e-133">**FormType**</span><span class="sxs-lookup"><span data-stu-id="f627e-133">**FormType**</span></span> | <span data-ttu-id="f627e-134">否（在 [ExtensionPoint](extensionpoint.md) 内），是（在 [OfficeApp](officeapp.md) 内）</span><span class="sxs-lookup"><span data-stu-id="f627e-134">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="f627e-p102">指定应用应出现在项目的读取还是编辑表单中。可以是以下值之一：`Read`、`Edit`、`ReadOrEdit`。如果在 `ExtensionPoint` 中的 `Rule` 上指定，则该值必须为 `Read`。</span><span class="sxs-lookup"><span data-stu-id="f627e-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="f627e-138">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="f627e-138">**ItemClass**</span></span> | <span data-ttu-id="f627e-139">否</span><span class="sxs-lookup"><span data-stu-id="f627e-139">No</span></span> | <span data-ttu-id="f627e-p103">指定要匹配的自定义邮件类别。有关详细信息，请参阅[在 Outlook 中为特定邮件类别激活邮件外接程序](../../outlook/activation-rules.md)。</span><span class="sxs-lookup"><span data-stu-id="f627e-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](../../outlook/activation-rules.md).</span></span> |
| <span data-ttu-id="f627e-142">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="f627e-142">**IncludeSubClasses**</span></span> | <span data-ttu-id="f627e-143">否</span><span class="sxs-lookup"><span data-stu-id="f627e-143">No</span></span> | <span data-ttu-id="f627e-144">指定当项目是指定邮件类别的子类时，该规则的计算结果是否应为 true；默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="f627e-144">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="f627e-145">示例</span><span class="sxs-lookup"><span data-stu-id="f627e-145">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="f627e-146">ItemHasAttachment 规则</span><span class="sxs-lookup"><span data-stu-id="f627e-146">ItemHasAttachment rule</span></span>

<span data-ttu-id="f627e-147">定义一个当项目包含附件时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="f627e-147">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="f627e-148">示例</span><span class="sxs-lookup"><span data-stu-id="f627e-148">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="f627e-149">ItemHasKnownEntity 规则</span><span class="sxs-lookup"><span data-stu-id="f627e-149">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="f627e-150">定义一个当项目主题或正文中包含指定实体类型的文本时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="f627e-150">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="f627e-151">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-151">Attributes</span></span>

| <span data-ttu-id="f627e-152">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-152">Attribute</span></span> | <span data-ttu-id="f627e-153">必需</span><span class="sxs-lookup"><span data-stu-id="f627e-153">Required</span></span> | <span data-ttu-id="f627e-154">说明</span><span class="sxs-lookup"><span data-stu-id="f627e-154">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="f627e-155">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="f627e-155">**EntityType**</span></span> | <span data-ttu-id="f627e-156">是</span><span class="sxs-lookup"><span data-stu-id="f627e-156">Yes</span></span> | <span data-ttu-id="f627e-p104">指定若想规则计算结果为 true 而必须存在的实体类型。可以是以下值之一：`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress` 或 `Contact`。</span><span class="sxs-lookup"><span data-stu-id="f627e-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="f627e-159">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="f627e-159">**RegExFilter**</span></span> | <span data-ttu-id="f627e-160">否</span><span class="sxs-lookup"><span data-stu-id="f627e-160">No</span></span> | <span data-ttu-id="f627e-161">指定一个针对此实体运行以进行激活的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="f627e-161">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="f627e-162">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="f627e-162">**FilterName**</span></span> | <span data-ttu-id="f627e-163">否</span><span class="sxs-lookup"><span data-stu-id="f627e-163">No</span></span> | <span data-ttu-id="f627e-164">指定正则表达式筛选器的名称，以便随后能够在你的外接程序代码中引用该名称。</span><span class="sxs-lookup"><span data-stu-id="f627e-164">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="f627e-165">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="f627e-165">**IgnoreCase**</span></span> | <span data-ttu-id="f627e-166">否</span><span class="sxs-lookup"><span data-stu-id="f627e-166">No</span></span> | <span data-ttu-id="f627e-167">指定在匹配由 **RegExFilter** 属性指定的正则表达式时是否忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="f627e-167">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="f627e-168">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="f627e-168">**Highlight**</span></span> | <span data-ttu-id="f627e-169">否</span><span class="sxs-lookup"><span data-stu-id="f627e-169">No</span></span> | <span data-ttu-id="f627e-p105">**注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的实体。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="f627e-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="f627e-174">示例</span><span class="sxs-lookup"><span data-stu-id="f627e-174">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="f627e-175">ItemHasRegularExpressionMatch 规则</span><span class="sxs-lookup"><span data-stu-id="f627e-175">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="f627e-176">定义一个如果可在项目的指定属性中找到指定的正则表达式的匹配项，则计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="f627e-176">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="f627e-177">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-177">Attributes</span></span>

| <span data-ttu-id="f627e-178">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-178">Attribute</span></span> | <span data-ttu-id="f627e-179">必需</span><span class="sxs-lookup"><span data-stu-id="f627e-179">Required</span></span> | <span data-ttu-id="f627e-180">说明</span><span class="sxs-lookup"><span data-stu-id="f627e-180">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="f627e-181">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="f627e-181">**RegExName**</span></span> | <span data-ttu-id="f627e-182">是</span><span class="sxs-lookup"><span data-stu-id="f627e-182">Yes</span></span> | <span data-ttu-id="f627e-183">指定正则表达式的名称，以便你能够在外接程序的代码中引用该表达式。</span><span class="sxs-lookup"><span data-stu-id="f627e-183">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="f627e-184">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="f627e-184">**RegExValue**</span></span> | <span data-ttu-id="f627e-185">是</span><span class="sxs-lookup"><span data-stu-id="f627e-185">Yes</span></span> | <span data-ttu-id="f627e-186">指定将对其求值的正则表达式以确定是否应显示邮件外接程序。</span><span class="sxs-lookup"><span data-stu-id="f627e-186">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="f627e-187">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="f627e-187">**PropertyName**</span></span> | <span data-ttu-id="f627e-188">是</span><span class="sxs-lookup"><span data-stu-id="f627e-188">Yes</span></span> | <span data-ttu-id="f627e-p106">指定正则表达式进行计算所依据的属性名称。可以是下列类型之一：`Subject`、`BodyAsPlaintext`、`BodyAsHTML` 或 `SenderSMTPAddress`。</span><span class="sxs-lookup"><span data-stu-id="f627e-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="f627e-191">如果指定 `BodyAsHTML`，则 Outlook 只会在项目正文为 HTML 时应用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="f627e-191">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="f627e-192">否则，Outlook 将不会返回该正则表达式的匹配项。</span><span class="sxs-lookup"><span data-stu-id="f627e-192">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="f627e-193">如果指定 `BodyAsPlaintext`，则 Outlook 将始终对项目正文应用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="f627e-193">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="f627e-194">**注释：** 如果指定 **Rule** 元素的 **Highlight** 属性，则必须将 **PropertyName** 属性设为 `BodyAsPlaintext`。</span><span class="sxs-lookup"><span data-stu-id="f627e-194">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="f627e-195">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="f627e-195">**IgnoreCase**</span></span> | <span data-ttu-id="f627e-196">否</span><span class="sxs-lookup"><span data-stu-id="f627e-196">No</span></span> | <span data-ttu-id="f627e-197">指定在匹配由 **RegExName** 属性指定的正则表达式时是否忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="f627e-197">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="f627e-198">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="f627e-198">**Highlight**</span></span> | <span data-ttu-id="f627e-199">否</span><span class="sxs-lookup"><span data-stu-id="f627e-199">No</span></span> | <span data-ttu-id="f627e-200">指定客户端应如何突出显示匹配的文本。</span><span class="sxs-lookup"><span data-stu-id="f627e-200">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="f627e-201">此属性仅适用于 **ExtensionPoint** 元素内的 **Rule** 元素。</span><span class="sxs-lookup"><span data-stu-id="f627e-201">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="f627e-202">可以是以下值之一：`all` 或 `none`。</span><span class="sxs-lookup"><span data-stu-id="f627e-202">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="f627e-203">如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="f627e-203">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="f627e-204">**注释：** 如果指定 **Rule** 元素的 **Highlight** 属性，则必须将 **PropertyName** 属性设为 `BodyAsPlaintext`。</span><span class="sxs-lookup"><span data-stu-id="f627e-204">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="f627e-205">示例</span><span class="sxs-lookup"><span data-stu-id="f627e-205">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="f627e-206">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="f627e-206">RuleCollection</span></span>

<span data-ttu-id="f627e-207">定义一个规则集合以及在计算这些规则时要使用的逻辑运算符。</span><span class="sxs-lookup"><span data-stu-id="f627e-207">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="f627e-208">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-208">Attributes</span></span>

| <span data-ttu-id="f627e-209">属性</span><span class="sxs-lookup"><span data-stu-id="f627e-209">Attribute</span></span> | <span data-ttu-id="f627e-210">必需</span><span class="sxs-lookup"><span data-stu-id="f627e-210">Required</span></span> | <span data-ttu-id="f627e-211">说明</span><span class="sxs-lookup"><span data-stu-id="f627e-211">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="f627e-212">**Mode**</span><span class="sxs-lookup"><span data-stu-id="f627e-212">**Mode**</span></span> | <span data-ttu-id="f627e-213">是</span><span class="sxs-lookup"><span data-stu-id="f627e-213">Yes</span></span> | <span data-ttu-id="f627e-p109">指定在计算此规则集时要使用的逻辑运算符。可以是以下类型之一：`And` 或 `Or`。</span><span class="sxs-lookup"><span data-stu-id="f627e-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="f627e-216">示例</span><span class="sxs-lookup"><span data-stu-id="f627e-216">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="f627e-217">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f627e-217">See also</span></span>

- [<span data-ttu-id="f627e-218">Outlook 加载项的激活规则</span><span class="sxs-lookup"><span data-stu-id="f627e-218">Activation rules for Outlook add-ins</span></span>](../../outlook/activation-rules.md)
- [<span data-ttu-id="f627e-219">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="f627e-219">Match strings in an Outlook item as well-known entities</span></span>](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="f627e-220">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="f627e-220">Use regular expression activation rules to show an Outlook add-in</span></span>](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
