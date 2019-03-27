---
title: 清单文件中的 Rule 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 07037c43c111f735a7354a048066e4c4a88f7637
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871512"
---
# <a name="rule-element"></a><span data-ttu-id="7dbe6-102">Rule 元素</span><span class="sxs-lookup"><span data-stu-id="7dbe6-102">Rule element</span></span>

<span data-ttu-id="7dbe6-103">指定应针对此上下文邮件加载项计算的激活规则。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-103">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="7dbe6-104">**加载项类型：** 邮件上下文加载项</span><span class="sxs-lookup"><span data-stu-id="7dbe6-104">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="7dbe6-105">包含于</span><span class="sxs-lookup"><span data-stu-id="7dbe6-105">Contained in</span></span>

- [<span data-ttu-id="7dbe6-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="7dbe6-106">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="7dbe6-107">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="7dbe6-107">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="7dbe6-108">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-108">Attributes</span></span>

| <span data-ttu-id="7dbe6-109">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-109">Attribute</span></span> | <span data-ttu-id="7dbe6-110">必需</span><span class="sxs-lookup"><span data-stu-id="7dbe6-110">Required</span></span> | <span data-ttu-id="7dbe6-111">说明</span><span class="sxs-lookup"><span data-stu-id="7dbe6-111">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7dbe6-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-112">**xsi:type**</span></span> | <span data-ttu-id="7dbe6-113">是</span><span class="sxs-lookup"><span data-stu-id="7dbe6-113">Yes</span></span> | <span data-ttu-id="7dbe6-114">正在定义的规则类型。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-114">The type of rule being defined.</span></span> |

<span data-ttu-id="7dbe6-115">此规则类型可以是下列类型之一。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-115">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="7dbe6-116">ItemIs</span><span class="sxs-lookup"><span data-stu-id="7dbe6-116">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="7dbe6-117">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="7dbe6-117">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="7dbe6-118">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="7dbe6-118">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="7dbe6-119">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="7dbe6-119">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="7dbe6-120">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="7dbe6-120">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="7dbe6-121">ItemIs 规则</span><span class="sxs-lookup"><span data-stu-id="7dbe6-121">ItemIs rule</span></span>

<span data-ttu-id="7dbe6-122">定义一个在选定项为指定类型时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-122">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="7dbe6-123">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-123">Attributes</span></span>

| <span data-ttu-id="7dbe6-124">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-124">Attribute</span></span> | <span data-ttu-id="7dbe6-125">必需</span><span class="sxs-lookup"><span data-stu-id="7dbe6-125">Required</span></span> | <span data-ttu-id="7dbe6-126">说明</span><span class="sxs-lookup"><span data-stu-id="7dbe6-126">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7dbe6-127">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-127">**ItemType**</span></span> | <span data-ttu-id="7dbe6-128">是</span><span class="sxs-lookup"><span data-stu-id="7dbe6-128">Yes</span></span> | <span data-ttu-id="7dbe6-p101">指定要匹配的项目类型。可以是 `Message` 或 `Appointment`。`Message` 项目类型包括电子邮件、会议请求、会议响应和会议取消。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="7dbe6-132">**FormType**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-132">**FormType**</span></span> | <span data-ttu-id="7dbe6-133">否（在 [ExtensionPoint](extensionpoint.md) 内），是（在 [OfficeApp](officeapp.md) 内）</span><span class="sxs-lookup"><span data-stu-id="7dbe6-133">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="7dbe6-p102">指定应用应出现在项目的读取还是编辑表单中。可以是以下值之一：`Read`、`Edit`、`ReadOrEdit`。如果在 `ExtensionPoint` 中的 `Rule` 上指定，则该值必须为 `Read`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="7dbe6-137">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-137">**ItemClass**</span></span> | <span data-ttu-id="7dbe6-138">否</span><span class="sxs-lookup"><span data-stu-id="7dbe6-138">No</span></span> | <span data-ttu-id="7dbe6-p103">指定要匹配的自定义邮件类别。有关详细信息，请参阅[在 Outlook 中为特定邮件类别激活邮件外接程序](/outlook/add-ins/activation-rules)。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="7dbe6-141">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-141">**IncludeSubClasses**</span></span> | <span data-ttu-id="7dbe6-142">否</span><span class="sxs-lookup"><span data-stu-id="7dbe6-142">No</span></span> | <span data-ttu-id="7dbe6-143">指定当项目是指定邮件类别的子类时，该规则的计算结果是否应为 true；默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-143">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="7dbe6-144">示例</span><span class="sxs-lookup"><span data-stu-id="7dbe6-144">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="7dbe6-145">ItemHasAttachment 规则</span><span class="sxs-lookup"><span data-stu-id="7dbe6-145">ItemHasAttachment rule</span></span>

<span data-ttu-id="7dbe6-146">定义一个当项目包含附件时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-146">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="7dbe6-147">示例</span><span class="sxs-lookup"><span data-stu-id="7dbe6-147">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="7dbe6-148">ItemHasKnownEntity 规则</span><span class="sxs-lookup"><span data-stu-id="7dbe6-148">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="7dbe6-149">定义一个当项目主题或正文中包含指定实体类型的文本时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-149">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="7dbe6-150">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-150">Attributes</span></span>

| <span data-ttu-id="7dbe6-151">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-151">Attribute</span></span> | <span data-ttu-id="7dbe6-152">必需</span><span class="sxs-lookup"><span data-stu-id="7dbe6-152">Required</span></span> | <span data-ttu-id="7dbe6-153">说明</span><span class="sxs-lookup"><span data-stu-id="7dbe6-153">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7dbe6-154">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-154">**EntityType**</span></span> | <span data-ttu-id="7dbe6-155">是</span><span class="sxs-lookup"><span data-stu-id="7dbe6-155">Yes</span></span> | <span data-ttu-id="7dbe6-p104">指定若想规则计算结果为 true 而必须存在的实体类型。可以是以下值之一：`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress` 或 `Contact`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="7dbe6-158">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-158">**RegExFilter**</span></span> | <span data-ttu-id="7dbe6-159">否</span><span class="sxs-lookup"><span data-stu-id="7dbe6-159">No</span></span> | <span data-ttu-id="7dbe6-160">指定一个针对此实体运行以进行激活的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-160">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="7dbe6-161">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-161">**FilterName**</span></span> | <span data-ttu-id="7dbe6-162">否</span><span class="sxs-lookup"><span data-stu-id="7dbe6-162">No</span></span> | <span data-ttu-id="7dbe6-163">指定正则表达式筛选器的名称，以便随后能够在你的外接程序代码中引用该名称。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-163">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="7dbe6-164">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-164">**IgnoreCase**</span></span> | <span data-ttu-id="7dbe6-165">否</span><span class="sxs-lookup"><span data-stu-id="7dbe6-165">No</span></span> | <span data-ttu-id="7dbe6-166">指定在匹配由 **RegExFilter** 属性指定的正则表达式时是否忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-166">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="7dbe6-167">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-167">**Highlight**</span></span> | <span data-ttu-id="7dbe6-168">否</span><span class="sxs-lookup"><span data-stu-id="7dbe6-168">No</span></span> | <span data-ttu-id="7dbe6-p105">**注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的实体。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="7dbe6-173">示例</span><span class="sxs-lookup"><span data-stu-id="7dbe6-173">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="7dbe6-174">ItemHasRegularExpressionMatch 规则</span><span class="sxs-lookup"><span data-stu-id="7dbe6-174">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="7dbe6-175">定义一个如果可在项目的指定属性中找到指定的正则表达式的匹配项，则计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-175">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="7dbe6-176">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-176">Attributes</span></span>

| <span data-ttu-id="7dbe6-177">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-177">Attribute</span></span> | <span data-ttu-id="7dbe6-178">必需</span><span class="sxs-lookup"><span data-stu-id="7dbe6-178">Required</span></span> | <span data-ttu-id="7dbe6-179">说明</span><span class="sxs-lookup"><span data-stu-id="7dbe6-179">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7dbe6-180">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-180">**RegExName**</span></span> | <span data-ttu-id="7dbe6-181">是</span><span class="sxs-lookup"><span data-stu-id="7dbe6-181">Yes</span></span> | <span data-ttu-id="7dbe6-182">指定正则表达式的名称，以便你能够在外接程序的代码中引用该表达式。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-182">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="7dbe6-183">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-183">**RegExValue**</span></span> | <span data-ttu-id="7dbe6-184">是</span><span class="sxs-lookup"><span data-stu-id="7dbe6-184">Yes</span></span> | <span data-ttu-id="7dbe6-185">指定将对其求值的正则表达式以确定是否应显示邮件外接程序。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-185">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="7dbe6-186">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-186">**PropertyName**</span></span> | <span data-ttu-id="7dbe6-187">是</span><span class="sxs-lookup"><span data-stu-id="7dbe6-187">Yes</span></span> | <span data-ttu-id="7dbe6-p106">指定正则表达式进行计算所依据的属性名称。可以是下列类型之一：`Subject`、`BodyAsPlaintext`、`BodyAsHTML` 或 `SenderSMTPAddress`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="7dbe6-190">如果指定 `BodyAsHTML`，则 Outlook 只会在项目正文为 HTML 时应用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-190">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="7dbe6-191">否则，Outlook 将不会返回该正则表达式的匹配项。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-191">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="7dbe6-192">如果指定 `BodyAsPlaintext`，则 Outlook 将始终对项目正文应用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-192">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="7dbe6-193">**注释：** 如果指定 **Rule** 元素的 **Highlight** 属性，则必须将 **PropertyName** 属性设为 `BodyAsPlaintext`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-193">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="7dbe6-194">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-194">**IgnoreCase**</span></span> | <span data-ttu-id="7dbe6-195">否</span><span class="sxs-lookup"><span data-stu-id="7dbe6-195">No</span></span> | <span data-ttu-id="7dbe6-196">指定在匹配由 **RegExName** 属性指定的正则表达式时是否忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-196">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="7dbe6-197">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-197">**Highlight**</span></span> | <span data-ttu-id="7dbe6-198">否</span><span class="sxs-lookup"><span data-stu-id="7dbe6-198">No</span></span> | <span data-ttu-id="7dbe6-199">指定客户端应如何突出显示匹配的文本。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-199">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="7dbe6-200">此属性仅适用于 **ExtensionPoint** 元素内的 **Rule** 元素。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-200">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="7dbe6-201">可以是以下值之一：`all` 或 `none`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-201">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="7dbe6-202">如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-202">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="7dbe6-203">**注释：** 如果指定 **Rule** 元素的 **Highlight** 属性，则必须将 **PropertyName** 属性设为 `BodyAsPlaintext`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-203">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="7dbe6-204">示例</span><span class="sxs-lookup"><span data-stu-id="7dbe6-204">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="7dbe6-205">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="7dbe6-205">RuleCollection</span></span>

<span data-ttu-id="7dbe6-206">定义一个规则集合以及在计算这些规则时要使用的逻辑运算符。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-206">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="7dbe6-207">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-207">Attributes</span></span>

| <span data-ttu-id="7dbe6-208">属性</span><span class="sxs-lookup"><span data-stu-id="7dbe6-208">Attribute</span></span> | <span data-ttu-id="7dbe6-209">必需</span><span class="sxs-lookup"><span data-stu-id="7dbe6-209">Required</span></span> | <span data-ttu-id="7dbe6-210">说明</span><span class="sxs-lookup"><span data-stu-id="7dbe6-210">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="7dbe6-211">**Mode**</span><span class="sxs-lookup"><span data-stu-id="7dbe6-211">**Mode**</span></span> | <span data-ttu-id="7dbe6-212">是</span><span class="sxs-lookup"><span data-stu-id="7dbe6-212">Yes</span></span> | <span data-ttu-id="7dbe6-p109">指定在计算此规则集时要使用的逻辑运算符。可以是以下类型之一：`And` 或 `Or`。</span><span class="sxs-lookup"><span data-stu-id="7dbe6-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="7dbe6-215">示例</span><span class="sxs-lookup"><span data-stu-id="7dbe6-215">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="7dbe6-216">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7dbe6-216">See also</span></span>

- [<span data-ttu-id="7dbe6-217">Outlook 加载项的激活规则</span><span class="sxs-lookup"><span data-stu-id="7dbe6-217">Activation rules for Outlook add-ins</span></span>](/outlook/add-ins/activation-rules)
- [<span data-ttu-id="7dbe6-218">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="7dbe6-218">Match strings in an Outlook item as well-known entities</span></span>](/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="7dbe6-219">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="7dbe6-219">Use regular expression activation rules to show an Outlook add-in</span></span>](/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)
