# <a name="rule-element"></a><span data-ttu-id="3e4a1-101">Rule 元素</span><span class="sxs-lookup"><span data-stu-id="3e4a1-101">Rule element</span></span>

<span data-ttu-id="3e4a1-102">指定应针对此上下文邮件加载项计算的激活规则。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-102">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="3e4a1-103">**加载项类型：** 邮件上下文加载项</span><span class="sxs-lookup"><span data-stu-id="3e4a1-103">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="3e4a1-104">包含于</span><span class="sxs-lookup"><span data-stu-id="3e4a1-104">Contained in</span></span>

- [<span data-ttu-id="3e4a1-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3e4a1-105">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="3e4a1-106">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="3e4a1-106">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="3e4a1-107">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-107">Attributes</span></span>

| <span data-ttu-id="3e4a1-108">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-108">Attribute</span></span> | <span data-ttu-id="3e4a1-109">必需</span><span class="sxs-lookup"><span data-stu-id="3e4a1-109">Required</span></span> | <span data-ttu-id="3e4a1-110">说明</span><span class="sxs-lookup"><span data-stu-id="3e4a1-110">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="3e4a1-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-111">**xsi:type**</span></span> | <span data-ttu-id="3e4a1-112">是</span><span class="sxs-lookup"><span data-stu-id="3e4a1-112">Yes</span></span> | <span data-ttu-id="3e4a1-113">正在定义的规则类型。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-113">The type of rule being defined.</span></span> |

<span data-ttu-id="3e4a1-114">此规则类型可以是下列类型之一。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-114">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="3e4a1-115">ItemIs</span><span class="sxs-lookup"><span data-stu-id="3e4a1-115">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="3e4a1-116">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="3e4a1-116">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="3e4a1-117">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="3e4a1-117">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="3e4a1-118">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="3e4a1-118">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="3e4a1-119">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="3e4a1-119">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="3e4a1-120">ItemIs 规则</span><span class="sxs-lookup"><span data-stu-id="3e4a1-120">ItemIs rule</span></span>

<span data-ttu-id="3e4a1-121">定义一个在选定项为指定类型时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-121">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="3e4a1-122">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-122">Attributes</span></span>

| <span data-ttu-id="3e4a1-123">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-123">Attribute</span></span> | <span data-ttu-id="3e4a1-124">必需</span><span class="sxs-lookup"><span data-stu-id="3e4a1-124">Required</span></span> | <span data-ttu-id="3e4a1-125">说明</span><span class="sxs-lookup"><span data-stu-id="3e4a1-125">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="3e4a1-126">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-126">**ItemType**</span></span> | <span data-ttu-id="3e4a1-127">是</span><span class="sxs-lookup"><span data-stu-id="3e4a1-127">Yes</span></span> | <span data-ttu-id="3e4a1-p101">指定要匹配的项目类型。可以是 `Message` 或 `Appointment`。`Message` 项目类型包括电子邮件、会议请求、会议响应和会议取消。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="3e4a1-131">**FormType**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-131">**FormType**</span></span> | <span data-ttu-id="3e4a1-132">否（在 [ExtensionPoint](extensionpoint.md) 内），是（在 [OfficeApp](officeapp.md) 内）</span><span class="sxs-lookup"><span data-stu-id="3e4a1-132">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="3e4a1-p102">指定应用应出现在项目的读取还是编辑表单中。可以是以下值之一：`Read`、`Edit`、`ReadOrEdit`。如果在 `ExtensionPoint` 中的 `Rule` 上指定，则该值必须为 `Read`。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="3e4a1-136">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-136">**ItemClass**</span></span> | <span data-ttu-id="3e4a1-137">否</span><span class="sxs-lookup"><span data-stu-id="3e4a1-137">No</span></span> | <span data-ttu-id="3e4a1-p103">指定要匹配的自定义邮件类别。有关详细信息，请参阅[在 Outlook 中为特定邮件类别激活邮件加载项](https://docs.microsoft.com/outlook/add-ins/activation-rules)。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="3e4a1-140">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-140">**IncludeSubClasses**</span></span> | <span data-ttu-id="3e4a1-141">否</span><span class="sxs-lookup"><span data-stu-id="3e4a1-141">No</span></span> | <span data-ttu-id="3e4a1-142">指定当项目是指定邮件类别的子类时，该规则的计算结果是否应为 true；默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-142">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="3e4a1-143">示例</span><span class="sxs-lookup"><span data-stu-id="3e4a1-143">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="3e4a1-144">ItemHasAttachment 规则</span><span class="sxs-lookup"><span data-stu-id="3e4a1-144">ItemHasAttachment rule</span></span>

<span data-ttu-id="3e4a1-145">定义一个当项目包含附件时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-145">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="3e4a1-146">示例</span><span class="sxs-lookup"><span data-stu-id="3e4a1-146">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="3e4a1-147">ItemHasKnownEntity 规则</span><span class="sxs-lookup"><span data-stu-id="3e4a1-147">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="3e4a1-148">定义一个当项目主题或正文中包含指定实体类型的文本时计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-148">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="3e4a1-149">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-149">Attributes</span></span>

| <span data-ttu-id="3e4a1-150">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-150">Attribute</span></span> | <span data-ttu-id="3e4a1-151">必需</span><span class="sxs-lookup"><span data-stu-id="3e4a1-151">Required</span></span> | <span data-ttu-id="3e4a1-152">说明</span><span class="sxs-lookup"><span data-stu-id="3e4a1-152">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="3e4a1-153">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-153">**EntityType**</span></span> | <span data-ttu-id="3e4a1-154">是</span><span class="sxs-lookup"><span data-stu-id="3e4a1-154">Yes</span></span> | <span data-ttu-id="3e4a1-p104">指定若想规则计算结果为 true 而必须存在的实体类型。可以为以下类型之一：`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress` 或 `Contact`。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="3e4a1-157">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-157">**RegExFilter**</span></span> | <span data-ttu-id="3e4a1-158">否</span><span class="sxs-lookup"><span data-stu-id="3e4a1-158">No</span></span> | <span data-ttu-id="3e4a1-159">指定一个针对此实体运行以进行激活的正则表达式。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-159">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="3e4a1-160">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-160">**FilterName**</span></span> | <span data-ttu-id="3e4a1-161">否</span><span class="sxs-lookup"><span data-stu-id="3e4a1-161">No</span></span> | <span data-ttu-id="3e4a1-162">指定正则表达式筛选器的名称，以便随后能够在你的外接程序代码中引用该名称。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-162">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="3e4a1-163">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-163">**IgnoreCase**</span></span> | <span data-ttu-id="3e4a1-164">否</span><span class="sxs-lookup"><span data-stu-id="3e4a1-164">No</span></span> | <span data-ttu-id="3e4a1-165">指定在运行由 **RegExFilter** 属性指定的正则表达式时忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-165">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="3e4a1-166">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-166">**Highlight**</span></span> | <span data-ttu-id="3e4a1-167">否</span><span class="sxs-lookup"><span data-stu-id="3e4a1-167">No</span></span> | <span data-ttu-id="3e4a1-p105">**注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的实体。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="3e4a1-172">示例</span><span class="sxs-lookup"><span data-stu-id="3e4a1-172">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="3e4a1-173">ItemHasRegularExpressionMatch 规则</span><span class="sxs-lookup"><span data-stu-id="3e4a1-173">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="3e4a1-174">定义一个如果可在项目的指定属性中找到指定的正则表达式的匹配项，则计算结果为 true 的规则。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-174">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="3e4a1-175">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-175">Attributes</span></span>

| <span data-ttu-id="3e4a1-176">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-176">Attribute</span></span> | <span data-ttu-id="3e4a1-177">必需</span><span class="sxs-lookup"><span data-stu-id="3e4a1-177">Required</span></span> | <span data-ttu-id="3e4a1-178">说明</span><span class="sxs-lookup"><span data-stu-id="3e4a1-178">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="3e4a1-179">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-179">**RegExName**</span></span> | <span data-ttu-id="3e4a1-180">是</span><span class="sxs-lookup"><span data-stu-id="3e4a1-180">Yes</span></span> | <span data-ttu-id="3e4a1-181">指定正则表达式的名称，以便你能够在外接程序的代码中引用该表达式。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-181">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="3e4a1-182">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-182">**RegExValue**</span></span> | <span data-ttu-id="3e4a1-183">是</span><span class="sxs-lookup"><span data-stu-id="3e4a1-183">Yes</span></span> | <span data-ttu-id="3e4a1-184">指定将对其求值的正则表达式以确定是否应显示邮件外接程序。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-184">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="3e4a1-185">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-185">**PropertyName**</span></span> | <span data-ttu-id="3e4a1-186">是</span><span class="sxs-lookup"><span data-stu-id="3e4a1-186">Yes</span></span> | <span data-ttu-id="3e4a1-p106">指定正则表达式进行计算所依据的属性名称。可以是下列类型之一：`Subject`、`BodyAsPlaintext`、`BodyAsHTML` 或 `SenderSTMPAddress`。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSTMPAddress`.</span></span> |
| <span data-ttu-id="3e4a1-189">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-189">**IgnoreCase**</span></span> | <span data-ttu-id="3e4a1-190">否</span><span class="sxs-lookup"><span data-stu-id="3e4a1-190">No</span></span> | <span data-ttu-id="3e4a1-191">指定在执行正则表达式时忽略大小写。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-191">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="3e4a1-192">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-192">**Highlight**</span></span> | <span data-ttu-id="3e4a1-193">否</span><span class="sxs-lookup"><span data-stu-id="3e4a1-193">No</span></span> | <span data-ttu-id="3e4a1-p107">**注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的文本。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="3e4a1-198">示例</span><span class="sxs-lookup"><span data-stu-id="3e4a1-198">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="3e4a1-199">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="3e4a1-199">RuleCollection</span></span>

<span data-ttu-id="3e4a1-200">定义一个规则集合以及在计算这些规则时要使用的逻辑运算符。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-200">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="3e4a1-201">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-201">Attributes</span></span>

| <span data-ttu-id="3e4a1-202">属性</span><span class="sxs-lookup"><span data-stu-id="3e4a1-202">Attribute</span></span> | <span data-ttu-id="3e4a1-203">必需</span><span class="sxs-lookup"><span data-stu-id="3e4a1-203">Required</span></span> | <span data-ttu-id="3e4a1-204">说明</span><span class="sxs-lookup"><span data-stu-id="3e4a1-204">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="3e4a1-205">**Mode**</span><span class="sxs-lookup"><span data-stu-id="3e4a1-205">**Mode**</span></span> | <span data-ttu-id="3e4a1-206">是</span><span class="sxs-lookup"><span data-stu-id="3e4a1-206">Yes</span></span> | <span data-ttu-id="3e4a1-p108">指定在计算此规则集时要使用的逻辑运算符。可以是以下类型之一：`And` 或 `Or`。</span><span class="sxs-lookup"><span data-stu-id="3e4a1-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="3e4a1-209">示例</span><span class="sxs-lookup"><span data-stu-id="3e4a1-209">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="3e4a1-210">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3e4a1-210">See also</span></span>

- [<span data-ttu-id="3e4a1-211">Outlook 加载项的激活规则</span><span class="sxs-lookup"><span data-stu-id="3e4a1-211">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="3e4a1-212">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="3e4a1-212">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="3e4a1-213">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="3e4a1-213">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)