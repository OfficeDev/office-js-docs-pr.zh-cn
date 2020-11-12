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
# <a name="override-element"></a><span data-ttu-id="62824-103">Override 元素</span><span class="sxs-lookup"><span data-stu-id="62824-103">Override element</span></span>

<span data-ttu-id="62824-104">提供一种方法，用于根据指定的条件重写清单设置的值。</span><span class="sxs-lookup"><span data-stu-id="62824-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="62824-105">有两种条件：</span><span class="sxs-lookup"><span data-stu-id="62824-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="62824-106">不同于默认的 Office 区域设置。</span><span class="sxs-lookup"><span data-stu-id="62824-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="62824-107">要求集支持的模式与默认模式不同。</span><span class="sxs-lookup"><span data-stu-id="62824-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="62824-108">有两种类型的 `<Override>` 元素，一个用于区域设置重写（称为 **LocaleTokenOverride** ），另一个用于要求集重写（称为 " **RequirementTokenOverride** "）。</span><span class="sxs-lookup"><span data-stu-id="62824-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride** , and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="62824-109">但没有 `type` 该元素的参数 `<Override>` 。</span><span class="sxs-lookup"><span data-stu-id="62824-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="62824-110">区别由父元素和父元素的类型确定。</span><span class="sxs-lookup"><span data-stu-id="62824-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="62824-111">`<Override>`元素中的元素， `<Token>` 其 `xsi:type` `RequirementToken` 类型必须为 **RequirementTokenOverride** 。</span><span class="sxs-lookup"><span data-stu-id="62824-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="62824-112">`<Override>`任何其他父元素中或类型元素内的元素 `<Override>` `LocaleToken` 都必须为 **LocaleTokenOverride** 类型。</span><span class="sxs-lookup"><span data-stu-id="62824-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="62824-113">以下各节分别介绍了每种类型。</span><span class="sxs-lookup"><span data-stu-id="62824-113">Each type is described in separate sections below.</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="62824-114">LocaleTokenOverride 类型的重写元素</span><span class="sxs-lookup"><span data-stu-id="62824-114">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="62824-115">`<Override>`元素表示条件，可读取为 "If ..."然后 ... "语句.</span><span class="sxs-lookup"><span data-stu-id="62824-115">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="62824-116">如果 `<Override>` 元素的类型为 **LocaleTokenOverride** ，则该 `Locale` 属性为条件， `Value` 属性随后会随后。</span><span class="sxs-lookup"><span data-stu-id="62824-116">If the `<Override>` element is of type **LocaleTokenOverride** , then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="62824-117">例如，以下是 "如果 Office 区域设置为 fr-fr"，则显示名称为 "Lecteur vidéo"。</span><span class="sxs-lookup"><span data-stu-id="62824-117">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="62824-118">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="62824-118">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="62824-119">语法</span><span class="sxs-lookup"><span data-stu-id="62824-119">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="62824-120">包含于</span><span class="sxs-lookup"><span data-stu-id="62824-120">Contained in</span></span>

|<span data-ttu-id="62824-121">元素</span><span class="sxs-lookup"><span data-stu-id="62824-121">Element</span></span>|
|:-----|
|[<span data-ttu-id="62824-122">CitationText</span><span class="sxs-lookup"><span data-stu-id="62824-122">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="62824-123">说明</span><span class="sxs-lookup"><span data-stu-id="62824-123">Description</span></span>](description.md)|
|[<span data-ttu-id="62824-124">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="62824-124">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="62824-125">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="62824-125">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="62824-126">DisplayName</span><span class="sxs-lookup"><span data-stu-id="62824-126">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="62824-127">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="62824-127">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="62824-128">IconUrl</span><span class="sxs-lookup"><span data-stu-id="62824-128">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="62824-129">QueryUri</span><span class="sxs-lookup"><span data-stu-id="62824-129">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="62824-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="62824-130">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="62824-131">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="62824-131">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="62824-132">标记</span><span class="sxs-lookup"><span data-stu-id="62824-132">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="62824-133">属性</span><span class="sxs-lookup"><span data-stu-id="62824-133">Attributes</span></span>

|<span data-ttu-id="62824-134">属性</span><span class="sxs-lookup"><span data-stu-id="62824-134">Attribute</span></span>|<span data-ttu-id="62824-135">类型</span><span class="sxs-lookup"><span data-stu-id="62824-135">Type</span></span>|<span data-ttu-id="62824-136">必需</span><span class="sxs-lookup"><span data-stu-id="62824-136">Required</span></span>|<span data-ttu-id="62824-137">说明</span><span class="sxs-lookup"><span data-stu-id="62824-137">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="62824-138">区域设置</span><span class="sxs-lookup"><span data-stu-id="62824-138">Locale</span></span>|<span data-ttu-id="62824-139">字符串</span><span class="sxs-lookup"><span data-stu-id="62824-139">string</span></span>|<span data-ttu-id="62824-140">必需</span><span class="sxs-lookup"><span data-stu-id="62824-140">required</span></span>|<span data-ttu-id="62824-141">为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。</span><span class="sxs-lookup"><span data-stu-id="62824-141">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="62824-142">值</span><span class="sxs-lookup"><span data-stu-id="62824-142">Value</span></span>|<span data-ttu-id="62824-143">字符串</span><span class="sxs-lookup"><span data-stu-id="62824-143">string</span></span>|<span data-ttu-id="62824-144">必需</span><span class="sxs-lookup"><span data-stu-id="62824-144">required</span></span>|<span data-ttu-id="62824-145">指定表示为指定区域设置的设置的值。</span><span class="sxs-lookup"><span data-stu-id="62824-145">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="62824-146">示例</span><span class="sxs-lookup"><span data-stu-id="62824-146">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="62824-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="62824-147">See also</span></span>

- [<span data-ttu-id="62824-148">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="62824-148">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="62824-149">键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="62824-149">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="62824-150">RequirementTokenOverride 类型的重写元素</span><span class="sxs-lookup"><span data-stu-id="62824-150">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="62824-151">`<Override>`元素表示条件，可读取为 "If ..."然后 ... "语句.</span><span class="sxs-lookup"><span data-stu-id="62824-151">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="62824-152">如果 `<Override>` 元素的类型为 **RequirementTokenOverride** ，则该子 `<Requirements>` 元素表示条件， `Value` 属性随后会随后。</span><span class="sxs-lookup"><span data-stu-id="62824-152">If the `<Override>` element is of type **RequirementTokenOverride** , then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="62824-153">例如，以下中的第一个 `<Override>` 是 "如果当前平台支持 FeatureOne 版本 1.7"，然后使用字符串 "oldAddinVersion" 替换 `${token.requirements}` 祖父 (的 URL 中的标记， `<ExtendedOverrides>` 而不是默认字符串 "upgrade" ) "。"</span><span class="sxs-lookup"><span data-stu-id="62824-153">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="62824-154">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="62824-154">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="62824-155">语法</span><span class="sxs-lookup"><span data-stu-id="62824-155">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="62824-156">包含于</span><span class="sxs-lookup"><span data-stu-id="62824-156">Contained in</span></span>

|<span data-ttu-id="62824-157">元素</span><span class="sxs-lookup"><span data-stu-id="62824-157">Element</span></span>|
|:-----|
|[<span data-ttu-id="62824-158">标记</span><span class="sxs-lookup"><span data-stu-id="62824-158">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="62824-159">必须包含</span><span class="sxs-lookup"><span data-stu-id="62824-159">Must contain</span></span>

|<span data-ttu-id="62824-160">元素</span><span class="sxs-lookup"><span data-stu-id="62824-160">Element</span></span>|<span data-ttu-id="62824-161">内容</span><span class="sxs-lookup"><span data-stu-id="62824-161">Content</span></span>|<span data-ttu-id="62824-162">邮件</span><span class="sxs-lookup"><span data-stu-id="62824-162">Mail</span></span>|<span data-ttu-id="62824-163">任务窗格</span><span class="sxs-lookup"><span data-stu-id="62824-163">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="62824-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="62824-164">Requirements</span></span>](requirements.md)|||<span data-ttu-id="62824-165">x</span><span class="sxs-lookup"><span data-stu-id="62824-165">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="62824-166">属性</span><span class="sxs-lookup"><span data-stu-id="62824-166">Attributes</span></span>

|<span data-ttu-id="62824-167">属性</span><span class="sxs-lookup"><span data-stu-id="62824-167">Attribute</span></span>|<span data-ttu-id="62824-168">类型</span><span class="sxs-lookup"><span data-stu-id="62824-168">Type</span></span>|<span data-ttu-id="62824-169">必需</span><span class="sxs-lookup"><span data-stu-id="62824-169">Required</span></span>|<span data-ttu-id="62824-170">说明</span><span class="sxs-lookup"><span data-stu-id="62824-170">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="62824-171">值</span><span class="sxs-lookup"><span data-stu-id="62824-171">Value</span></span>|<span data-ttu-id="62824-172">字符串</span><span class="sxs-lookup"><span data-stu-id="62824-172">string</span></span>|<span data-ttu-id="62824-173">必需</span><span class="sxs-lookup"><span data-stu-id="62824-173">required</span></span>|<span data-ttu-id="62824-174">满足条件时的祖父令牌的值。</span><span class="sxs-lookup"><span data-stu-id="62824-174">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="62824-175">示例</span><span class="sxs-lookup"><span data-stu-id="62824-175">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="62824-176">另请参阅</span><span class="sxs-lookup"><span data-stu-id="62824-176">See also</span></span>

- [<span data-ttu-id="62824-177">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="62824-177">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="62824-178">在清单中设置 Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="62824-178">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="62824-179">键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="62824-179">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
