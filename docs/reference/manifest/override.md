---
title: 清单文件中的 Override 元素
description: Override 元素使您能够根据指定的条件指定设置的值。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 2c66503f9f95155a096b1b6fb23332eed8422da6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996310"
---
# <a name="override-element"></a>Override 元素

提供一种方法，用于根据指定的条件重写清单设置的值。 有两种条件：

- 不同于默认的 Office 区域设置。
- 要求集支持的模式与默认模式不同。

有两种类型的 `<Override>` 元素，一个用于区域设置重写（称为 **LocaleTokenOverride** ），另一个用于要求集重写（称为 " **RequirementTokenOverride** "）。 但没有 `type` 该元素的参数 `<Override>` 。 区别由父元素和父元素的类型确定。 `<Override>`元素中的元素， `<Token>` 其 `xsi:type` `RequirementToken` 类型必须为 **RequirementTokenOverride** 。 `<Override>`任何其他父元素中或类型元素内的元素 `<Override>` `LocaleToken` 都必须为 **LocaleTokenOverride** 类型。 以下各节分别介绍了每种类型。

## <a name="override-element-of-type-localetokenoverride"></a>LocaleTokenOverride 类型的重写元素

`<Override>`元素表示条件，可读取为 "If ..."然后 ... "语句. 如果 `<Override>` 元素的类型为 **LocaleTokenOverride** ，则该 `Locale` 属性为条件， `Value` 属性随后会随后。 例如，以下是 "如果 Office 区域设置为 fr-fr"，则显示名称为 "Lecteur vidéo"。

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

## <a name="override-element-of-type-requirementtokenoverride"></a>RequirementTokenOverride 类型的重写元素

`<Override>`元素表示条件，可读取为 "If ..."然后 ... "语句. 如果 `<Override>` 元素的类型为 **RequirementTokenOverride** ，则该子 `<Requirements>` 元素表示条件， `Value` 属性随后会随后。 例如，以下中的第一个 `<Override>` 是 "如果当前平台支持 FeatureOne 版本 1.7"，然后使用字符串 "oldAddinVersion" 替换 `${token.requirements}` 祖父 (的 URL 中的标记， `<ExtendedOverrides>` 而不是默认字符串 "upgrade" ) "。"

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
|值|字符串|必需|满足条件时的祖父令牌的值。|

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
