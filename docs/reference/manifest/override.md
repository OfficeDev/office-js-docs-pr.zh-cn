---
title: 清单文件中的 Override 元素
description: Override 元素使您能够根据指定条件指定设置的值。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: d2146cc1f44e829bc78076c8093b2ebf791dc722
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505337"
---
# <a name="override-element"></a><span data-ttu-id="15cdc-103">Override 元素</span><span class="sxs-lookup"><span data-stu-id="15cdc-103">Override element</span></span>

<span data-ttu-id="15cdc-104">提供一种根据指定条件替代清单设置的值的方法。</span><span class="sxs-lookup"><span data-stu-id="15cdc-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="15cdc-105">有两种类型的条件：</span><span class="sxs-lookup"><span data-stu-id="15cdc-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="15cdc-106">不同于默认值的 Office 区域设置。</span><span class="sxs-lookup"><span data-stu-id="15cdc-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="15cdc-107">与默认模式不同的要求集支持模式。</span><span class="sxs-lookup"><span data-stu-id="15cdc-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="15cdc-108">有两种类型的元素，一种用于区域设置重写，称为 `<Override>` **LocaleTokenOverride，** 另一种用于要求集替代，称为 **RequirementTokenOverride。**</span><span class="sxs-lookup"><span data-stu-id="15cdc-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride**, and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="15cdc-109">但元素 `type` 没有 `<Override>` 参数。</span><span class="sxs-lookup"><span data-stu-id="15cdc-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="15cdc-110">差异由父元素和父元素的类型确定。</span><span class="sxs-lookup"><span data-stu-id="15cdc-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="15cdc-111">元素 `<Override>` 位于其类型为 `<Token>` `xsi:type` `RequirementToken` **RequirementTokenOverride** 的元素内。</span><span class="sxs-lookup"><span data-stu-id="15cdc-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="15cdc-112">任何其他 `<Override>` 父元素内或类型元素内的元素必须为 `<Override>` `LocaleToken` **LocaleTokenOverride 类型**。</span><span class="sxs-lookup"><span data-stu-id="15cdc-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="15cdc-113">以下各节分别介绍了每种类型。</span><span class="sxs-lookup"><span data-stu-id="15cdc-113">Each type is described in separate sections below.</span></span> <span data-ttu-id="15cdc-114">有关当此元素是元素的子级时使用此元素的信息，请参阅"处理清单 `<Token>` [的扩展重写"。](../../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="15cdc-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="15cdc-115">LocaleTokenOverride 类型的 Override 元素</span><span class="sxs-lookup"><span data-stu-id="15cdc-115">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="15cdc-116">元素 `<Override>` 表示条件，并可以读取为"If ...then ..."语句。</span><span class="sxs-lookup"><span data-stu-id="15cdc-116">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="15cdc-117">如果 `<Override>` 元素的类型为 **LocaleTokenOverride，** 则该属性为 `Locale` 条件，而 `Value` 该属性是结果。</span><span class="sxs-lookup"><span data-stu-id="15cdc-117">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="15cdc-118">例如，下面的内容为"如果 Office 区域设置为 fr-fr，则显示名称为"Lecteur vidéo"。</span><span class="sxs-lookup"><span data-stu-id="15cdc-118">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="15cdc-119">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="15cdc-119">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="15cdc-120">语法</span><span class="sxs-lookup"><span data-stu-id="15cdc-120">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="15cdc-121">包含于</span><span class="sxs-lookup"><span data-stu-id="15cdc-121">Contained in</span></span>

|<span data-ttu-id="15cdc-122">元素</span><span class="sxs-lookup"><span data-stu-id="15cdc-122">Element</span></span>|
|:-----|
|[<span data-ttu-id="15cdc-123">CitationText</span><span class="sxs-lookup"><span data-stu-id="15cdc-123">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="15cdc-124">说明</span><span class="sxs-lookup"><span data-stu-id="15cdc-124">Description</span></span>](description.md)|
|[<span data-ttu-id="15cdc-125">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="15cdc-125">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="15cdc-126">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="15cdc-126">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="15cdc-127">DisplayName</span><span class="sxs-lookup"><span data-stu-id="15cdc-127">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="15cdc-128">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="15cdc-128">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="15cdc-129">IconUrl</span><span class="sxs-lookup"><span data-stu-id="15cdc-129">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="15cdc-130">QueryUri</span><span class="sxs-lookup"><span data-stu-id="15cdc-130">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="15cdc-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="15cdc-131">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="15cdc-132">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="15cdc-132">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="15cdc-133">标记</span><span class="sxs-lookup"><span data-stu-id="15cdc-133">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="15cdc-134">属性</span><span class="sxs-lookup"><span data-stu-id="15cdc-134">Attributes</span></span>

|<span data-ttu-id="15cdc-135">属性</span><span class="sxs-lookup"><span data-stu-id="15cdc-135">Attribute</span></span>|<span data-ttu-id="15cdc-136">类型</span><span class="sxs-lookup"><span data-stu-id="15cdc-136">Type</span></span>|<span data-ttu-id="15cdc-137">必需</span><span class="sxs-lookup"><span data-stu-id="15cdc-137">Required</span></span>|<span data-ttu-id="15cdc-138">说明</span><span class="sxs-lookup"><span data-stu-id="15cdc-138">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="15cdc-139">区域设置</span><span class="sxs-lookup"><span data-stu-id="15cdc-139">Locale</span></span>|<span data-ttu-id="15cdc-140">字符串</span><span class="sxs-lookup"><span data-stu-id="15cdc-140">string</span></span>|<span data-ttu-id="15cdc-141">必需</span><span class="sxs-lookup"><span data-stu-id="15cdc-141">required</span></span>|<span data-ttu-id="15cdc-142">为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。</span><span class="sxs-lookup"><span data-stu-id="15cdc-142">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="15cdc-143">值</span><span class="sxs-lookup"><span data-stu-id="15cdc-143">Value</span></span>|<span data-ttu-id="15cdc-144">字符串</span><span class="sxs-lookup"><span data-stu-id="15cdc-144">string</span></span>|<span data-ttu-id="15cdc-145">必需</span><span class="sxs-lookup"><span data-stu-id="15cdc-145">required</span></span>|<span data-ttu-id="15cdc-146">指定表示为指定区域设置的设置的值。</span><span class="sxs-lookup"><span data-stu-id="15cdc-146">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="15cdc-147">示例</span><span class="sxs-lookup"><span data-stu-id="15cdc-147">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="15cdc-148">另请参阅</span><span class="sxs-lookup"><span data-stu-id="15cdc-148">See also</span></span>

- [<span data-ttu-id="15cdc-149">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="15cdc-149">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="15cdc-150">键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="15cdc-150">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="15cdc-151">RequirementTokenOverride 类型的 Override 元素</span><span class="sxs-lookup"><span data-stu-id="15cdc-151">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="15cdc-152">元素 `<Override>` 表示条件，并可以读取为"If ...then ..."语句。</span><span class="sxs-lookup"><span data-stu-id="15cdc-152">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="15cdc-153">如果 `<Override>` 元素的类型 **为 RequirementTokenOverride，** 则子元素表示条件，而 `<Requirements>` `Value` 该属性是结果。</span><span class="sxs-lookup"><span data-stu-id="15cdc-153">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="15cdc-154">例如，下面的第一个内容是"如果当前平台支持 `<Override>` FeatureOne 版本 1.7，则使用字符串"oldAddinVersion"代替 (的 URL 中的令牌，而不是默认字符串 `${token.requirements}` `<ExtendedOverrides>` "upgrade") "。</span><span class="sxs-lookup"><span data-stu-id="15cdc-154">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="15cdc-155">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="15cdc-155">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="15cdc-156">语法</span><span class="sxs-lookup"><span data-stu-id="15cdc-156">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="15cdc-157">包含于</span><span class="sxs-lookup"><span data-stu-id="15cdc-157">Contained in</span></span>

|<span data-ttu-id="15cdc-158">元素</span><span class="sxs-lookup"><span data-stu-id="15cdc-158">Element</span></span>|
|:-----|
|[<span data-ttu-id="15cdc-159">标记</span><span class="sxs-lookup"><span data-stu-id="15cdc-159">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="15cdc-160">必须包含</span><span class="sxs-lookup"><span data-stu-id="15cdc-160">Must contain</span></span>

|<span data-ttu-id="15cdc-161">元素</span><span class="sxs-lookup"><span data-stu-id="15cdc-161">Element</span></span>|<span data-ttu-id="15cdc-162">内容</span><span class="sxs-lookup"><span data-stu-id="15cdc-162">Content</span></span>|<span data-ttu-id="15cdc-163">邮件</span><span class="sxs-lookup"><span data-stu-id="15cdc-163">Mail</span></span>|<span data-ttu-id="15cdc-164">任务窗格</span><span class="sxs-lookup"><span data-stu-id="15cdc-164">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="15cdc-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="15cdc-165">Requirements</span></span>](requirements.md)|||<span data-ttu-id="15cdc-166">x</span><span class="sxs-lookup"><span data-stu-id="15cdc-166">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="15cdc-167">属性</span><span class="sxs-lookup"><span data-stu-id="15cdc-167">Attributes</span></span>

|<span data-ttu-id="15cdc-168">属性</span><span class="sxs-lookup"><span data-stu-id="15cdc-168">Attribute</span></span>|<span data-ttu-id="15cdc-169">类型</span><span class="sxs-lookup"><span data-stu-id="15cdc-169">Type</span></span>|<span data-ttu-id="15cdc-170">必需</span><span class="sxs-lookup"><span data-stu-id="15cdc-170">Required</span></span>|<span data-ttu-id="15cdc-171">说明</span><span class="sxs-lookup"><span data-stu-id="15cdc-171">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="15cdc-172">值</span><span class="sxs-lookup"><span data-stu-id="15cdc-172">Value</span></span>|<span data-ttu-id="15cdc-173">字符串</span><span class="sxs-lookup"><span data-stu-id="15cdc-173">string</span></span>|<span data-ttu-id="15cdc-174">必需</span><span class="sxs-lookup"><span data-stu-id="15cdc-174">required</span></span>|<span data-ttu-id="15cdc-175">满足条件时令牌的值。</span><span class="sxs-lookup"><span data-stu-id="15cdc-175">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="15cdc-176">示例</span><span class="sxs-lookup"><span data-stu-id="15cdc-176">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="15cdc-177">另请参阅</span><span class="sxs-lookup"><span data-stu-id="15cdc-177">See also</span></span>

- [<span data-ttu-id="15cdc-178">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="15cdc-178">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="15cdc-179">在清单中设置 Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="15cdc-179">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="15cdc-180">键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="15cdc-180">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
