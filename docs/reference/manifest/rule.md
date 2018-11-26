# <a name="rule-element"></a>Rule 元素

指定应针对此上下文邮件加载项计算的激活规则。

**加载项类型：** 邮件上下文加载项

## <a name="contained-in"></a>包含于

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md)

## <a name="attributes"></a>属性

| 属性 | 必需 | 说明 |
|:-----|:-----|:-----|
| **xsi:type** | 是 | 正在定义的规则类型。 |

此规则类型可以是下列类型之一。

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)
- [RuleCollection](#rulecollection)

## <a name="itemis-rule"></a>ItemIs 规则

定义一个在选定项为指定类型时计算结果为 true 的规则。

### <a name="attributes"></a>属性

| 属性 | 必需 | 说明 |
|:-----|:-----|:-----|
| **ItemType** | 是 | 指定要匹配的项目类型。可以是 `Message` 或 `Appointment`。`Message` 项目类型包括电子邮件、会议请求、会议响应和会议取消。 |
| **FormType** | 否（在 [ExtensionPoint](extensionpoint.md) 内），是（在 [OfficeApp](officeapp.md) 内） | 指定应用应出现在项目的读取还是编辑表单中。可以是以下值之一：`Read`、`Edit`、`ReadOrEdit`。如果在 `ExtensionPoint` 中的 `Rule` 上指定，则该值必须为 `Read`。 |
| **ItemClass** | 否 | 指定要匹配的自定义邮件类别。有关详细信息，请参阅[在 Outlook 中为特定邮件类别激活邮件加载项](https://docs.microsoft.com/outlook/add-ins/activation-rules)。 |
| **IncludeSubClasses** | 否 | 指定当项目是指定邮件类别的子类时，该规则的计算结果是否应为 true；默认值为 `false`。 |

### <a name="example"></a>示例

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a>ItemHasAttachment 规则

定义一个当项目包含附件时计算结果为 true 的规则。

### <a name="example"></a>示例

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity 规则

定义一个当项目主题或正文中包含指定实体类型的文本时计算结果为 true 的规则。

### <a name="attributes"></a>属性

| 属性 | 必需 | 说明 |
|:-----|:-----|:-----|
| **EntityType** | 是 | 指定若想规则计算结果为 true 而必须存在的实体类型。可以为以下类型之一：`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress` 或 `Contact`。 |
| **RegExFilter** | 否 | 指定一个针对此实体运行以进行激活的正则表达式。 |
| **FilterName** | 否 | 指定正则表达式筛选器的名称，以便随后能够在你的外接程序代码中引用该名称。 |
| **IgnoreCase** | 否 | 指定在运行由 **RegExFilter** 属性指定的正则表达式时忽略大小写。 |
| **Highlight** | 否 | **注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的实体。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。 |

### <a name="example"></a>示例

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch 规则

定义一个如果可在项目的指定属性中找到指定的正则表达式的匹配项，则计算结果为 true 的规则。

### <a name="attributes"></a>属性

| 属性 | 必需 | 说明 |
|:-----|:-----|:-----|
| **RegExName** | 是 | 指定正则表达式的名称，以便你能够在外接程序的代码中引用该表达式。 |
| **RegExValue** | 是 | 指定将对其求值的正则表达式以确定是否应显示邮件外接程序。 |
| **PropertyName** | 是 | 指定正则表达式进行计算所依据的属性名称。可以是下列类型之一：`Subject`、`BodyAsPlaintext`、`BodyAsHTML` 或 `SenderSTMPAddress`。 |
| **IgnoreCase** | 否 | 指定在执行正则表达式时忽略大小写。 |
| **Highlight** | 否 | **注意：** 这仅适用于 **ExtensionPoint** 元素中的 **Rule** 元素。指定客户端应如何突出显示匹配的文本。可以是以下值之一：`all` 或 `none`。如果未指定，则默认值为 `all`。 |

### <a name="example"></a>示例

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

定义一个规则集合以及在计算这些规则时要使用的逻辑运算符。

### <a name="attributes"></a>属性

| 属性 | 必需 | 说明 |
|:-----|:-----|:-----|
| **Mode** | 是 | 指定在计算此规则集时要使用的逻辑运算符。可以是以下类型之一：`And` 或 `Or`。 |

### <a name="example"></a>示例

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>另请参阅

- [Outlook 加载项的激活规则](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [将 Outlook 项中的字符串作为已知实体进行匹配](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [使用正则表达式激活规则显示 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)