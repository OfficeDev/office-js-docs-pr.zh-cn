---
title: 清单文件中的 Rule 元素
description: Rule 元素指定应为此上下文邮件外接程序计算的激活规则。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: c4094cdf9e9006bbc49d180cb79845527461a543
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608109"
---
# <a name="rule-element"></a><span data-ttu-id="dedd4-103">Rule 元素</span><span class="sxs-lookup"><span data-stu-id="dedd4-103">Rule element</span></span>

<span data-ttu-id="dedd4-104">指定应针对此上下文邮件外接程序计算的激活规则。</span><span class="sxs-lookup"><span data-stu-id="dedd4-104">Specifies the activation rules that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="dedd4-105">**外接类型：** 邮件（上下文）</span><span class="sxs-lookup"><span data-stu-id="dedd4-105">**Add-in type:** Mail (contextual)</span></span>

## <a name="contained-in"></a><span data-ttu-id="dedd4-106">包含于</span><span class="sxs-lookup"><span data-stu-id="dedd4-106">Contained in</span></span>

- [<span data-ttu-id="dedd4-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="dedd4-107">OfficeApp</span></span>](officeapp.md)
- <span data-ttu-id="dedd4-108">[ExtensionPoint](extensionpoint.md) （[**CustomPane** （已弃用）](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)， [**DetectedEntity**](extensionpoint.md#detectedentity)）</span><span class="sxs-lookup"><span data-stu-id="dedd4-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span></span>

## <a name="attributes"></a><span data-ttu-id="dedd4-109">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-109">Attributes</span></span>

| <span data-ttu-id="dedd4-110">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-110">Attribute</span></span> | <span data-ttu-id="dedd4-111">必需</span><span class="sxs-lookup"><span data-stu-id="dedd4-111">Required</span></span> | <span data-ttu-id="dedd4-112">说明</span><span class="sxs-lookup"><span data-stu-id="dedd4-112">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="dedd4-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="dedd4-113">**xsi:type**</span></span> | <span data-ttu-id="dedd4-114">是</span><span class="sxs-lookup"><span data-stu-id="dedd4-114">Yes</span></span> | <span data-ttu-id="dedd4-115">正在定义的规则类型。</span><span class="sxs-lookup"><span data-stu-id="dedd4-115">The type of rule being defined.</span></span> |

<span data-ttu-id="dedd4-116">此规则类型可以是下列类型之一。</span><span class="sxs-lookup"><span data-stu-id="dedd4-116">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="dedd4-117">ItemIs</span><span class="sxs-lookup"><span data-stu-id="dedd4-117">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="dedd4-118">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="dedd4-118">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="dedd4-119">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="dedd4-119">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="dedd4-120">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="dedd4-120">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="dedd4-121">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="dedd4-121">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="dedd4-122">ItemIs 规则</span><span class="sxs-lookup"><span data-stu-id="dedd4-122">ItemIs rule</span></span>

<span data-ttu-id="dedd4-123">定义一个在选定项为指定类型时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="dedd4-123">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="dedd4-124">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-124">Attributes</span></span>

| <span data-ttu-id="dedd4-125">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-125">Attribute</span></span> | <span data-ttu-id="dedd4-126">必需</span><span class="sxs-lookup"><span data-stu-id="dedd4-126">Required</span></span> | <span data-ttu-id="dedd4-127">Description</span><span class="sxs-lookup"><span data-stu-id="dedd4-127">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="dedd4-128">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="dedd4-128">**ItemType**</span></span> | <span data-ttu-id="dedd4-129">是</span><span class="sxs-lookup"><span data-stu-id="dedd4-129">Yes</span></span> | <span data-ttu-id="dedd4-p101">指定要匹配的项目类型。可以是 `Message` 或 `Appointment`。`Message` 项目类型包括电子邮件、会议请求、会议响应和会议取消。</span><span class="sxs-lookup"><span data-stu-id="dedd4-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="dedd4-133">**FormType**</span><span class="sxs-lookup"><span data-stu-id="dedd4-133">**FormType**</span></span> | <span data-ttu-id="dedd4-134">否（在 [ExtensionPoint](extensionpoint.md) 内），是（在 [OfficeApp](officeapp.md) 内）</span><span class="sxs-lookup"><span data-stu-id="dedd4-134">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="dedd4-p102">指定应用应出现在项目的读取还是编辑表单中。可以是以下值之一：`Read`、`Edit`、`ReadOrEdit`。如果在 `ExtensionPoint` 中的 `Rule` 上指定，则该值必须为 `Read`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="dedd4-138">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="dedd4-138">**ItemClass**</span></span> | <span data-ttu-id="dedd4-139">否</span><span class="sxs-lookup"><span data-stu-id="dedd4-139">No</span></span> | <span data-ttu-id="dedd4-p103">指定要匹配的自定义邮件类别。有关详细信息，请参阅[在 Outlook 中为特定邮件类别激活邮件外接程序](../../outlook/activation-rules.md)。</span><span class="sxs-lookup"><span data-stu-id="dedd4-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](../../outlook/activation-rules.md).</span></span> |
| <span data-ttu-id="dedd4-142">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="dedd4-142">**IncludeSubClasses**</span></span> | <span data-ttu-id="dedd4-143">否</span><span class="sxs-lookup"><span data-stu-id="dedd4-143">No</span></span> | <span data-ttu-id="dedd4-144">指定当项目是指定邮件类别的子类时，该规则的计算结果是否应为 true；默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-144">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="dedd4-145">示例</span><span class="sxs-lookup"><span data-stu-id="dedd4-145">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="dedd4-146">ItemHasAttachment 规则</span><span class="sxs-lookup"><span data-stu-id="dedd4-146">ItemHasAttachment rule</span></span>

<span data-ttu-id="dedd4-147">定义一个当项目包含附件时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="dedd4-147">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="dedd4-148">示例</span><span class="sxs-lookup"><span data-stu-id="dedd4-148">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="dedd4-149">ItemHasKnownEntity 规则</span><span class="sxs-lookup"><span data-stu-id="dedd4-149">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="dedd4-150">定义一个当项目主题或正文中包含指定实体类型的文本时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="dedd4-150">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="dedd4-151">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-151">Attributes</span></span>

| <span data-ttu-id="dedd4-152">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-152">Attribute</span></span> | <span data-ttu-id="dedd4-153">必需</span><span class="sxs-lookup"><span data-stu-id="dedd4-153">Required</span></span> | <span data-ttu-id="dedd4-154">Description</span><span class="sxs-lookup"><span data-stu-id="dedd4-154">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="dedd4-155">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="dedd4-155">**EntityType**</span></span> | <span data-ttu-id="dedd4-156">是</span><span class="sxs-lookup"><span data-stu-id="dedd4-156">Yes</span></span> | <span data-ttu-id="dedd4-p104">指定若想规则计算结果为 true 而必须存在的实体类型。可以是以下值之一：`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress` 或 `Contact`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="dedd4-159">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="dedd4-159">**RegExFilter**</span></span> | <span data-ttu-id="dedd4-160">否</span><span class="sxs-lookup"><span data-stu-id="dedd4-160">No</span></span> | <span data-ttu-id="dedd4-161">指定一个针对此实体运行以进行激活的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="dedd4-161">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="dedd4-162">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="dedd4-162">**FilterName**</span></span> | <span data-ttu-id="dedd4-163">否</span><span class="sxs-lookup"><span data-stu-id="dedd4-163">No</span></span> | <span data-ttu-id="dedd4-164">指定正则表达式筛选器的名称，以便随后能够在你的外接程序代码中引用该名称。</span><span class="sxs-lookup"><span data-stu-id="dedd4-164">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="dedd4-165">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="dedd4-165">**IgnoreCase**</span></span> | <span data-ttu-id="dedd4-166">否</span><span class="sxs-lookup"><span data-stu-id="dedd4-166">No</span></span> | <span data-ttu-id="dedd4-167">指定在匹配由 **RegExFilter** 属性指定的正则表达式时是否忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="dedd4-167">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="dedd4-168">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="dedd4-168">**Highlight**</span></span> | <span data-ttu-id="dedd4-169">否</span><span class="sxs-lookup"><span data-stu-id="dedd4-169">No</span></span> | <span data-ttu-id="dedd4-p105">**注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的实体。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="dedd4-174">示例</span><span class="sxs-lookup"><span data-stu-id="dedd4-174">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="dedd4-175">ItemHasRegularExpressionMatch 规则</span><span class="sxs-lookup"><span data-stu-id="dedd4-175">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="dedd4-176">定义一个如果可在项目的指定属性中找到指定的正则表达式的匹配项，则计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="dedd4-176">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="dedd4-177">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-177">Attributes</span></span>

| <span data-ttu-id="dedd4-178">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-178">Attribute</span></span> | <span data-ttu-id="dedd4-179">必需</span><span class="sxs-lookup"><span data-stu-id="dedd4-179">Required</span></span> | <span data-ttu-id="dedd4-180">Description</span><span class="sxs-lookup"><span data-stu-id="dedd4-180">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="dedd4-181">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="dedd4-181">**RegExName**</span></span> | <span data-ttu-id="dedd4-182">是</span><span class="sxs-lookup"><span data-stu-id="dedd4-182">Yes</span></span> | <span data-ttu-id="dedd4-183">指定正则表达式的名称，以便你能够在外接程序的代码中引用该表达式。</span><span class="sxs-lookup"><span data-stu-id="dedd4-183">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="dedd4-184">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="dedd4-184">**RegExValue**</span></span> | <span data-ttu-id="dedd4-185">是</span><span class="sxs-lookup"><span data-stu-id="dedd4-185">Yes</span></span> | <span data-ttu-id="dedd4-186">指定将对其求值的正则表达式以确定是否应显示邮件外接程序。</span><span class="sxs-lookup"><span data-stu-id="dedd4-186">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="dedd4-187">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="dedd4-187">**PropertyName**</span></span> | <span data-ttu-id="dedd4-188">是</span><span class="sxs-lookup"><span data-stu-id="dedd4-188">Yes</span></span> | <span data-ttu-id="dedd4-p106">指定正则表达式进行计算所依据的属性名称。可以是下列类型之一：`Subject`、`BodyAsPlaintext`、`BodyAsHTML` 或 `SenderSMTPAddress`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="dedd4-191">如果指定 `BodyAsHTML`，则 Outlook 只会在项目正文为 HTML 时应用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="dedd4-191">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="dedd4-192">否则，Outlook 将不会返回该正则表达式的匹配项。</span><span class="sxs-lookup"><span data-stu-id="dedd4-192">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="dedd4-193">如果指定 `BodyAsPlaintext`，则 Outlook 将始终对项目正文应用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="dedd4-193">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="dedd4-194">**注释：** 如果指定 **Rule** 元素的 **Highlight** 属性，则必须将 **PropertyName** 属性设为 `BodyAsPlaintext`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-194">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="dedd4-195">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="dedd4-195">**IgnoreCase**</span></span> | <span data-ttu-id="dedd4-196">否</span><span class="sxs-lookup"><span data-stu-id="dedd4-196">No</span></span> | <span data-ttu-id="dedd4-197">指定在匹配由 **RegExName** 属性指定的正则表达式时是否忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="dedd4-197">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="dedd4-198">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="dedd4-198">**Highlight**</span></span> | <span data-ttu-id="dedd4-199">否</span><span class="sxs-lookup"><span data-stu-id="dedd4-199">No</span></span> | <span data-ttu-id="dedd4-200">指定客户端应如何突出显示匹配的文本。</span><span class="sxs-lookup"><span data-stu-id="dedd4-200">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="dedd4-201">此属性仅适用于 **ExtensionPoint** 元素内的 **Rule** 元素。</span><span class="sxs-lookup"><span data-stu-id="dedd4-201">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="dedd4-202">可以是以下值之一：`all` 或 `none`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-202">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="dedd4-203">如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-203">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="dedd4-204">**注释：** 如果指定 **Rule** 元素的 **Highlight** 属性，则必须将 **PropertyName** 属性设为 `BodyAsPlaintext`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-204">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="dedd4-205">示例</span><span class="sxs-lookup"><span data-stu-id="dedd4-205">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="dedd4-206">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="dedd4-206">RuleCollection</span></span>

<span data-ttu-id="dedd4-207">定义一个规则集合以及在计算这些规则时要使用的逻辑运算符。</span><span class="sxs-lookup"><span data-stu-id="dedd4-207">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="dedd4-208">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-208">Attributes</span></span>

| <span data-ttu-id="dedd4-209">属性</span><span class="sxs-lookup"><span data-stu-id="dedd4-209">Attribute</span></span> | <span data-ttu-id="dedd4-210">必需</span><span class="sxs-lookup"><span data-stu-id="dedd4-210">Required</span></span> | <span data-ttu-id="dedd4-211">Description</span><span class="sxs-lookup"><span data-stu-id="dedd4-211">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="dedd4-212">**Mode**</span><span class="sxs-lookup"><span data-stu-id="dedd4-212">**Mode**</span></span> | <span data-ttu-id="dedd4-213">是</span><span class="sxs-lookup"><span data-stu-id="dedd4-213">Yes</span></span> | <span data-ttu-id="dedd4-p109">指定在计算此规则集时要使用的逻辑运算符。可以是以下类型之一：`And` 或 `Or`。</span><span class="sxs-lookup"><span data-stu-id="dedd4-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="dedd4-216">示例</span><span class="sxs-lookup"><span data-stu-id="dedd4-216">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="dedd4-217">另请参阅</span><span class="sxs-lookup"><span data-stu-id="dedd4-217">See also</span></span>

- [<span data-ttu-id="dedd4-218">Outlook 加载项的激活规则</span><span class="sxs-lookup"><span data-stu-id="dedd4-218">Activation rules for Outlook add-ins</span></span>](../../outlook/activation-rules.md)
- [<span data-ttu-id="dedd4-219">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="dedd4-219">Match strings in an Outlook item as well-known entities</span></span>](../../outlook/match-strings-in-an-item-as-well-known-entities.md)    
- [<span data-ttu-id="dedd4-220">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="dedd4-220">Use regular expression activation rules to show an Outlook add-in</span></span>](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
