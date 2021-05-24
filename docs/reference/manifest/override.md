---
title: 清单文件中的 Override 元素
description: Override 元素使您能够根据指定条件指定设置的值。
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd270fa19750810238b42c26c2abc35a61c1bac8
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590902"
---
# <a name="override-element"></a>Override 元素

提供一种根据指定条件替代清单设置的值的方法。 有三种类型的条件：

- 与Office区域设置不同的区域设置，称为 `LocaleToken` **LocaleTokenOverride**。
- 与默认模式不同的要求集支持模式，称为 `RequirementToken` **RequirementTokenOverride**。
- 源不同于默认的 ，称为 `Runtime` **RuntimeOverride**。

`<Override>`元素内的元素必须为 `<Runtime>` **RuntimeOverride 类型**。

元素 `overrideType` 没有 `<Override>` 属性。 差异由父元素和父元素的类型确定。 元素位于 其 为 的元素内，其类型 `<Override>` `<Token>` 必须为 `xsi:type` `RequirementToken` **RequirementTokenOverride**。 任何其他 `<Override>` 父元素内或类型元素内的元素必须为 `<Override>` `LocaleToken` **LocaleTokenOverride 类型**。 有关当此元素是元素的子元素时该元素的使用详细信息，请参阅使用清单 `<Token>` [的扩展替代](../../develop/extended-overrides.md)。

每种类型在本文稍后的单独部分中介绍。

## <a name="override-element-for-localetoken"></a>的 Override 元素 `LocaleToken`

元素 `<Override>` 表示条件，可读为"If ...then ..."语句。 如果 `<Override>` 元素的类型为 **LocaleTokenOverride**，则属性为条件 `Locale` ， `Value` 而 属性为结果。 例如，以下为"如果 Office区域设置是 fr-fr，则显示名称是"Lecteur vidéo"。

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

**加载项类型：** 内容、任务窗格和邮件

### <a name="syntax"></a>语法

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a>包含于

|元素|
|:-----|
|[CitationText](citationtext.md)|
|[说明](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|
|[标记](token.md)|

### <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|区域设置|字符串|必需|为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。|
|值|字符串|必需|指定表示为指定区域设置的设置的值。|

### <a name="examples"></a>示例

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
```

### <a name="see-also"></a>另请参阅

- [Office 外接程序的本地化](../../develop/localization.md)
- [键盘快捷方式](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a>的 Override 元素 `RequirementToken`

元素 `<Override>` 表示条件，可读为"If ...then ..."语句。 如果 `<Override>` 元素的类型为 **RequirementTokenOverride**，则子元素表示条件，而 `<Requirements>` `Value` 属性是结果。 例如，下面的第一个代码为"如果当前平台支持 `<Override>` FeatureOne 版本 1.7，则使用字符串'oldAddinVersion'代替 (而不是默认字符串 `${token.requirements}` `<ExtendedOverrides>` "upgrade") 。"

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

**外接程序类型：** 任务窗格

### <a name="syntax"></a>语法

```XML
<Override Value="string" />
```

### <a name="contained-in"></a>包含于

|元素|
|:-----|
|[标记](token.md)|

### <a name="must-contain"></a>必须包含

|元素|内容|邮件|任务窗格|
|:-----|:-----|:-----|:-----|
|[Requirements](requirements.md)|||x|

### <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|值|字符串|必需|满足条件时令牌的值。|

### <a name="example"></a>示例

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [在清单中设置 Requirements 元素](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [键盘快捷方式](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime"></a>的 Override 元素 `Runtime`

> [!IMPORTANT]
> 邮箱要求集 [1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) 中引入了对此元素的支持，该功能具有基于 [事件的激活功能](../../outlook/autolaunch.md)。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

元素 `<Override>` 表示条件，可读为"If ...then ..."语句。 如果 `<Override>` 元素的类型为 **RuntimeOverride**，则 属性为 `type` 条件， `resid` 属性为结果。 例如，以下代码为"如果类型为'javascript'，则 `resid` 为'JSRuntime.Url'"。Outlook桌面需要此元素用于[LaunchEvent 扩展点](../../reference/manifest/extensionpoint.md#launchevent)处理程序。

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

**外接程序类型：** 邮件

### <a name="syntax"></a>语法

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a>包含于

- [运行时](runtime.md)

### <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|**类型**|string|是|指定此替代的语言。 目前， `"javascript"` 是唯一受支持的选项。|
|**resid**|string|是|指定 JavaScript 文件的 URL 位置，该文件应替代在父 [Runtime](runtime.md) 元素 中定义的默认 HTML 的 URL 位置 `resid` 。 `resid`不能超过 32 个字符，并且必须与 元素中的 `id` `Url` 元素的 属性 `Resources` 匹配。|

### <a name="examples"></a>示例

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a>另请参阅

- [运行时](runtime.md)
- [配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md)
