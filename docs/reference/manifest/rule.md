---
title: 清单文件中的 Rule 元素
description: ''
ms.date: 11/30/2018
ms.openlocfilehash: ce7763ecb4ef81587ccacbd4090a6f412baf99b2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433110"
---
# <a name="rule-element"></a><span data-ttu-id="70fef-102">Rule 元素</span><span class="sxs-lookup"><span data-stu-id="70fef-102">Rule element</span></span>

<span data-ttu-id="70fef-103">指定应针对此上下文邮件加载项计算的激活规则。</span><span class="sxs-lookup"><span data-stu-id="70fef-103">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="70fef-104">**加载项类型：** 邮件上下文加载项</span><span class="sxs-lookup"><span data-stu-id="70fef-104">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="70fef-105">包含于</span><span class="sxs-lookup"><span data-stu-id="70fef-105">Contained in</span></span>

- [<span data-ttu-id="70fef-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="70fef-106">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="70fef-107">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="70fef-107">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="70fef-108">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-108">Attributes</span></span>

| <span data-ttu-id="70fef-109">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-109">Attribute</span></span> | <span data-ttu-id="70fef-110">必需</span><span class="sxs-lookup"><span data-stu-id="70fef-110">Required</span></span> | <span data-ttu-id="70fef-111">说明</span><span class="sxs-lookup"><span data-stu-id="70fef-111">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="70fef-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="70fef-112">**xsi:type**</span></span> | <span data-ttu-id="70fef-113">是</span><span class="sxs-lookup"><span data-stu-id="70fef-113">Yes</span></span> | <span data-ttu-id="70fef-114">正在定义的规则类型。</span><span class="sxs-lookup"><span data-stu-id="70fef-114">The type of rule being defined.</span></span> |

<span data-ttu-id="70fef-115">此规则类型可以是下列类型之一。</span><span class="sxs-lookup"><span data-stu-id="70fef-115">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="70fef-116">ItemIs</span><span class="sxs-lookup"><span data-stu-id="70fef-116">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="70fef-117">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="70fef-117">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="70fef-118">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="70fef-118">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="70fef-119">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="70fef-119">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="70fef-120">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="70fef-120">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="70fef-121">ItemIs 规则</span><span class="sxs-lookup"><span data-stu-id="70fef-121">ItemIs rule</span></span>

<span data-ttu-id="70fef-122">定义一个在选定项为指定类型时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="70fef-122">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="70fef-123">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-123">Attributes</span></span>

| <span data-ttu-id="70fef-124">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-124">Attribute</span></span> | <span data-ttu-id="70fef-125">必需</span><span class="sxs-lookup"><span data-stu-id="70fef-125">Required</span></span> | <span data-ttu-id="70fef-126">说明</span><span class="sxs-lookup"><span data-stu-id="70fef-126">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="70fef-127">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="70fef-127">**ItemType**</span></span> | <span data-ttu-id="70fef-128">是</span><span class="sxs-lookup"><span data-stu-id="70fef-128">Yes</span></span> | <span data-ttu-id="70fef-p101">指定要匹配的项目类型。可以是 `Message` 或 `Appointment`。`Message` 项目类型包括电子邮件、会议请求、会议响应和会议取消。</span><span class="sxs-lookup"><span data-stu-id="70fef-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="70fef-132">**FormType**</span><span class="sxs-lookup"><span data-stu-id="70fef-132">**FormType**</span></span> | <span data-ttu-id="70fef-133">否（在 [ExtensionPoint](extensionpoint.md) 内），是（在 [OfficeApp](officeapp.md) 内）</span><span class="sxs-lookup"><span data-stu-id="70fef-133">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="70fef-p102">指定应用应出现在项目的读取还是编辑表单中。可以是以下值之一：`Read`、`Edit`、`ReadOrEdit`。如果在 `ExtensionPoint` 中的 `Rule` 上指定，则该值必须为 `Read`。</span><span class="sxs-lookup"><span data-stu-id="70fef-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="70fef-137">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="70fef-137">**ItemClass**</span></span> | <span data-ttu-id="70fef-138">否</span><span class="sxs-lookup"><span data-stu-id="70fef-138">No</span></span> | <span data-ttu-id="70fef-p103">指定要匹配的自定义邮件类别。有关详细信息，请参阅[在 Outlook 中为特定邮件类别激活邮件加载项](https://docs.microsoft.com/outlook/add-ins/activation-rules)。</span><span class="sxs-lookup"><span data-stu-id="70fef-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="70fef-141">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="70fef-141">**IncludeSubClasses**</span></span> | <span data-ttu-id="70fef-142">否</span><span class="sxs-lookup"><span data-stu-id="70fef-142">No</span></span> | <span data-ttu-id="70fef-143">指定当项目是指定邮件类别的子类时，该规则的计算结果是否应为 true；默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="70fef-143">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="70fef-144">示例</span><span class="sxs-lookup"><span data-stu-id="70fef-144">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="70fef-145">ItemHasAttachment 规则</span><span class="sxs-lookup"><span data-stu-id="70fef-145">ItemHasAttachment rule</span></span>

<span data-ttu-id="70fef-146">定义一个当项目包含附件时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="70fef-146">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="70fef-147">示例</span><span class="sxs-lookup"><span data-stu-id="70fef-147">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="70fef-148">ItemHasKnownEntity 规则</span><span class="sxs-lookup"><span data-stu-id="70fef-148">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="70fef-149">定义一个当项目主题或正文中包含指定实体类型的文本时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="70fef-149">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="70fef-150">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-150">Attributes</span></span>

| <span data-ttu-id="70fef-151">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-151">Attribute</span></span> | <span data-ttu-id="70fef-152">必需</span><span class="sxs-lookup"><span data-stu-id="70fef-152">Required</span></span> | <span data-ttu-id="70fef-153">说明</span><span class="sxs-lookup"><span data-stu-id="70fef-153">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="70fef-154">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="70fef-154">**EntityType**</span></span> | <span data-ttu-id="70fef-155">是</span><span class="sxs-lookup"><span data-stu-id="70fef-155">Yes</span></span> | <span data-ttu-id="70fef-p104">指定若想规则计算结果为 true 而必须存在的实体类型。可以为以下类型之一：`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress` 或 `Contact`。</span><span class="sxs-lookup"><span data-stu-id="70fef-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="70fef-158">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="70fef-158">**RegExFilter**</span></span> | <span data-ttu-id="70fef-159">否</span><span class="sxs-lookup"><span data-stu-id="70fef-159">No</span></span> | <span data-ttu-id="70fef-160">指定一个针对此实体运行以进行激活的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="70fef-160">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="70fef-161">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="70fef-161">**FilterName**</span></span> | <span data-ttu-id="70fef-162">否</span><span class="sxs-lookup"><span data-stu-id="70fef-162">No</span></span> | <span data-ttu-id="70fef-163">指定正则表达式筛选器的名称，以便随后能够在你的外接程序代码中引用该名称。</span><span class="sxs-lookup"><span data-stu-id="70fef-163">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="70fef-164">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="70fef-164">**IgnoreCase**</span></span> | <span data-ttu-id="70fef-165">否</span><span class="sxs-lookup"><span data-stu-id="70fef-165">No</span></span> | <span data-ttu-id="70fef-166">指定在运行由 **RegExFilter** 属性指定的正则表达式时忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="70fef-166">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="70fef-167">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="70fef-167">**Highlight**</span></span> | <span data-ttu-id="70fef-168">否</span><span class="sxs-lookup"><span data-stu-id="70fef-168">No</span></span> | <span data-ttu-id="70fef-p105">**注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的实体。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="70fef-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="70fef-173">示例</span><span class="sxs-lookup"><span data-stu-id="70fef-173">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="70fef-174">ItemHasRegularExpressionMatch 规则</span><span class="sxs-lookup"><span data-stu-id="70fef-174">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="70fef-175">定义一个如果可在项目的指定属性中找到指定的正则表达式的匹配项，则计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="70fef-175">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="70fef-176">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-176">Attributes</span></span>

| <span data-ttu-id="70fef-177">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-177">Attribute</span></span> | <span data-ttu-id="70fef-178">必需</span><span class="sxs-lookup"><span data-stu-id="70fef-178">Required</span></span> | <span data-ttu-id="70fef-179">说明</span><span class="sxs-lookup"><span data-stu-id="70fef-179">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="70fef-180">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="70fef-180">**RegExName**</span></span> | <span data-ttu-id="70fef-181">是</span><span class="sxs-lookup"><span data-stu-id="70fef-181">Yes</span></span> | <span data-ttu-id="70fef-182">指定正则表达式的名称，以便你能够在外接程序的代码中引用该表达式。</span><span class="sxs-lookup"><span data-stu-id="70fef-182">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="70fef-183">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="70fef-183">**RegExValue**</span></span> | <span data-ttu-id="70fef-184">是</span><span class="sxs-lookup"><span data-stu-id="70fef-184">Yes</span></span> | <span data-ttu-id="70fef-185">指定将对其求值的正则表达式以确定是否应显示邮件外接程序。</span><span class="sxs-lookup"><span data-stu-id="70fef-185">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="70fef-186">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="70fef-186">**PropertyName**</span></span> | <span data-ttu-id="70fef-187">是</span><span class="sxs-lookup"><span data-stu-id="70fef-187">Yes</span></span> | <span data-ttu-id="70fef-p106">指定正则表达式进行计算所依据的属性名称。可以是下列类型之一：`Subject`、`BodyAsPlaintext`、`BodyAsHTML` 或 `SenderSMTPAddress`。</span><span class="sxs-lookup"><span data-stu-id="70fef-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span> |
| <span data-ttu-id="70fef-190">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="70fef-190">**IgnoreCase**</span></span> | <span data-ttu-id="70fef-191">否</span><span class="sxs-lookup"><span data-stu-id="70fef-191">No</span></span> | <span data-ttu-id="70fef-192">指定在执行正则表达式时忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="70fef-192">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="70fef-193">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="70fef-193">**Highlight**</span></span> | <span data-ttu-id="70fef-194">否</span><span class="sxs-lookup"><span data-stu-id="70fef-194">No</span></span> | <span data-ttu-id="70fef-p107">**注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的文本。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="70fef-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="70fef-199">示例</span><span class="sxs-lookup"><span data-stu-id="70fef-199">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="70fef-200">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="70fef-200">RuleCollection</span></span>

<span data-ttu-id="70fef-201">定义一个规则集合以及在计算这些规则时要使用的逻辑运算符。</span><span class="sxs-lookup"><span data-stu-id="70fef-201">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="70fef-202">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-202">Attributes</span></span>

| <span data-ttu-id="70fef-203">属性</span><span class="sxs-lookup"><span data-stu-id="70fef-203">Attribute</span></span> | <span data-ttu-id="70fef-204">必需</span><span class="sxs-lookup"><span data-stu-id="70fef-204">Required</span></span> | <span data-ttu-id="70fef-205">说明</span><span class="sxs-lookup"><span data-stu-id="70fef-205">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="70fef-206">**Mode**</span><span class="sxs-lookup"><span data-stu-id="70fef-206">**Mode**</span></span> | <span data-ttu-id="70fef-207">是</span><span class="sxs-lookup"><span data-stu-id="70fef-207">Yes</span></span> | <span data-ttu-id="70fef-p108">指定在计算此规则集时要使用的逻辑运算符。可以是以下类型之一：`And` 或 `Or`。</span><span class="sxs-lookup"><span data-stu-id="70fef-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="70fef-210">示例</span><span class="sxs-lookup"><span data-stu-id="70fef-210">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="70fef-211">另请参阅</span><span class="sxs-lookup"><span data-stu-id="70fef-211">See also</span></span>

- [<span data-ttu-id="70fef-212">Outlook 加载项的激活规则</span><span class="sxs-lookup"><span data-stu-id="70fef-212">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="70fef-213">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="70fef-213">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="70fef-214">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="70fef-214">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)