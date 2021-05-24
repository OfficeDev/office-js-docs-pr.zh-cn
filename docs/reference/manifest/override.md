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
# <a name="override-element"></a><span data-ttu-id="d538e-103">Override 元素</span><span class="sxs-lookup"><span data-stu-id="d538e-103">Override element</span></span>

<span data-ttu-id="d538e-104">提供一种根据指定条件替代清单设置的值的方法。</span><span class="sxs-lookup"><span data-stu-id="d538e-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="d538e-105">有三种类型的条件：</span><span class="sxs-lookup"><span data-stu-id="d538e-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="d538e-106">与Office区域设置不同的区域设置，称为 `LocaleToken` **LocaleTokenOverride**。</span><span class="sxs-lookup"><span data-stu-id="d538e-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="d538e-107">与默认模式不同的要求集支持模式，称为 `RequirementToken` **RequirementTokenOverride**。</span><span class="sxs-lookup"><span data-stu-id="d538e-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="d538e-108">源不同于默认的 ，称为 `Runtime` **RuntimeOverride**。</span><span class="sxs-lookup"><span data-stu-id="d538e-108">The source is different from the default `Runtime`, called **RuntimeOverride**.</span></span>

<span data-ttu-id="d538e-109">`<Override>`元素内的元素必须为 `<Runtime>` **RuntimeOverride 类型**。</span><span class="sxs-lookup"><span data-stu-id="d538e-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="d538e-110">元素 `overrideType` 没有 `<Override>` 属性。</span><span class="sxs-lookup"><span data-stu-id="d538e-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="d538e-111">差异由父元素和父元素的类型确定。</span><span class="sxs-lookup"><span data-stu-id="d538e-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="d538e-112">元素位于 其 为 的元素内，其类型 `<Override>` `<Token>` 必须为 `xsi:type` `RequirementToken` **RequirementTokenOverride**。</span><span class="sxs-lookup"><span data-stu-id="d538e-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="d538e-113">任何其他 `<Override>` 父元素内或类型元素内的元素必须为 `<Override>` `LocaleToken` **LocaleTokenOverride 类型**。</span><span class="sxs-lookup"><span data-stu-id="d538e-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="d538e-114">有关当此元素是元素的子元素时该元素的使用详细信息，请参阅使用清单 `<Token>` [的扩展替代](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="d538e-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="d538e-115">每种类型在本文稍后的单独部分中介绍。</span><span class="sxs-lookup"><span data-stu-id="d538e-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="d538e-116">的 Override 元素 `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="d538e-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="d538e-117">元素 `<Override>` 表示条件，可读为"If ...then ..."语句。</span><span class="sxs-lookup"><span data-stu-id="d538e-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="d538e-118">如果 `<Override>` 元素的类型为 **LocaleTokenOverride**，则属性为条件 `Locale` ， `Value` 而 属性为结果。</span><span class="sxs-lookup"><span data-stu-id="d538e-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="d538e-119">例如，以下为"如果 Office区域设置是 fr-fr，则显示名称是"Lecteur vidéo"。</span><span class="sxs-lookup"><span data-stu-id="d538e-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="d538e-120">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="d538e-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="d538e-121">语法</span><span class="sxs-lookup"><span data-stu-id="d538e-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="d538e-122">包含于</span><span class="sxs-lookup"><span data-stu-id="d538e-122">Contained in</span></span>

|<span data-ttu-id="d538e-123">元素</span><span class="sxs-lookup"><span data-stu-id="d538e-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="d538e-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="d538e-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="d538e-125">说明</span><span class="sxs-lookup"><span data-stu-id="d538e-125">Description</span></span>](description.md)|
|[<span data-ttu-id="d538e-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="d538e-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="d538e-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="d538e-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="d538e-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="d538e-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="d538e-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="d538e-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="d538e-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="d538e-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="d538e-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="d538e-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="d538e-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d538e-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="d538e-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="d538e-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="d538e-134">标记</span><span class="sxs-lookup"><span data-stu-id="d538e-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="d538e-135">属性</span><span class="sxs-lookup"><span data-stu-id="d538e-135">Attributes</span></span>

|<span data-ttu-id="d538e-136">属性</span><span class="sxs-lookup"><span data-stu-id="d538e-136">Attribute</span></span>|<span data-ttu-id="d538e-137">类型</span><span class="sxs-lookup"><span data-stu-id="d538e-137">Type</span></span>|<span data-ttu-id="d538e-138">必需</span><span class="sxs-lookup"><span data-stu-id="d538e-138">Required</span></span>|<span data-ttu-id="d538e-139">说明</span><span class="sxs-lookup"><span data-stu-id="d538e-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d538e-140">区域设置</span><span class="sxs-lookup"><span data-stu-id="d538e-140">Locale</span></span>|<span data-ttu-id="d538e-141">字符串</span><span class="sxs-lookup"><span data-stu-id="d538e-141">string</span></span>|<span data-ttu-id="d538e-142">必需</span><span class="sxs-lookup"><span data-stu-id="d538e-142">required</span></span>|<span data-ttu-id="d538e-143">为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。</span><span class="sxs-lookup"><span data-stu-id="d538e-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="d538e-144">值</span><span class="sxs-lookup"><span data-stu-id="d538e-144">Value</span></span>|<span data-ttu-id="d538e-145">字符串</span><span class="sxs-lookup"><span data-stu-id="d538e-145">string</span></span>|<span data-ttu-id="d538e-146">必需</span><span class="sxs-lookup"><span data-stu-id="d538e-146">required</span></span>|<span data-ttu-id="d538e-147">指定表示为指定区域设置的设置的值。</span><span class="sxs-lookup"><span data-stu-id="d538e-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="d538e-148">示例</span><span class="sxs-lookup"><span data-stu-id="d538e-148">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="d538e-149">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d538e-149">See also</span></span>

- [<span data-ttu-id="d538e-150">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="d538e-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="d538e-151">键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="d538e-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="d538e-152">的 Override 元素 `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="d538e-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="d538e-153">元素 `<Override>` 表示条件，可读为"If ...then ..."语句。</span><span class="sxs-lookup"><span data-stu-id="d538e-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="d538e-154">如果 `<Override>` 元素的类型为 **RequirementTokenOverride**，则子元素表示条件，而 `<Requirements>` `Value` 属性是结果。</span><span class="sxs-lookup"><span data-stu-id="d538e-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="d538e-155">例如，下面的第一个代码为"如果当前平台支持 `<Override>` FeatureOne 版本 1.7，则使用字符串'oldAddinVersion'代替 (而不是默认字符串 `${token.requirements}` `<ExtendedOverrides>` "upgrade") 。"</span><span class="sxs-lookup"><span data-stu-id="d538e-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="d538e-156">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d538e-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="d538e-157">语法</span><span class="sxs-lookup"><span data-stu-id="d538e-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="d538e-158">包含于</span><span class="sxs-lookup"><span data-stu-id="d538e-158">Contained in</span></span>

|<span data-ttu-id="d538e-159">元素</span><span class="sxs-lookup"><span data-stu-id="d538e-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="d538e-160">标记</span><span class="sxs-lookup"><span data-stu-id="d538e-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="d538e-161">必须包含</span><span class="sxs-lookup"><span data-stu-id="d538e-161">Must contain</span></span>

|<span data-ttu-id="d538e-162">元素</span><span class="sxs-lookup"><span data-stu-id="d538e-162">Element</span></span>|<span data-ttu-id="d538e-163">内容</span><span class="sxs-lookup"><span data-stu-id="d538e-163">Content</span></span>|<span data-ttu-id="d538e-164">邮件</span><span class="sxs-lookup"><span data-stu-id="d538e-164">Mail</span></span>|<span data-ttu-id="d538e-165">任务窗格</span><span class="sxs-lookup"><span data-stu-id="d538e-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="d538e-166">Requirements</span><span class="sxs-lookup"><span data-stu-id="d538e-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="d538e-167">x</span><span class="sxs-lookup"><span data-stu-id="d538e-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="d538e-168">属性</span><span class="sxs-lookup"><span data-stu-id="d538e-168">Attributes</span></span>

|<span data-ttu-id="d538e-169">属性</span><span class="sxs-lookup"><span data-stu-id="d538e-169">Attribute</span></span>|<span data-ttu-id="d538e-170">类型</span><span class="sxs-lookup"><span data-stu-id="d538e-170">Type</span></span>|<span data-ttu-id="d538e-171">必需</span><span class="sxs-lookup"><span data-stu-id="d538e-171">Required</span></span>|<span data-ttu-id="d538e-172">说明</span><span class="sxs-lookup"><span data-stu-id="d538e-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d538e-173">值</span><span class="sxs-lookup"><span data-stu-id="d538e-173">Value</span></span>|<span data-ttu-id="d538e-174">字符串</span><span class="sxs-lookup"><span data-stu-id="d538e-174">string</span></span>|<span data-ttu-id="d538e-175">必需</span><span class="sxs-lookup"><span data-stu-id="d538e-175">required</span></span>|<span data-ttu-id="d538e-176">满足条件时令牌的值。</span><span class="sxs-lookup"><span data-stu-id="d538e-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="d538e-177">示例</span><span class="sxs-lookup"><span data-stu-id="d538e-177">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="d538e-178">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d538e-178">See also</span></span>

- [<span data-ttu-id="d538e-179">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="d538e-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d538e-180">在清单中设置 Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="d538e-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="d538e-181">键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="d538e-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime"></a><span data-ttu-id="d538e-182">的 Override 元素 `Runtime`</span><span class="sxs-lookup"><span data-stu-id="d538e-182">Override element for `Runtime`</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d538e-183">邮箱要求集 [1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) 中引入了对此元素的支持，该功能具有基于 [事件的激活功能](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="d538e-183">Support for this element was introduced in [Mailbox requirement set 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) with the [event-based activation feature](../../outlook/autolaunch.md).</span></span> <span data-ttu-id="d538e-184">请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="d538e-184">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="d538e-185">元素 `<Override>` 表示条件，可读为"If ...then ..."语句。</span><span class="sxs-lookup"><span data-stu-id="d538e-185">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="d538e-186">如果 `<Override>` 元素的类型为 **RuntimeOverride**，则 属性为 `type` 条件， `resid` 属性为结果。</span><span class="sxs-lookup"><span data-stu-id="d538e-186">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="d538e-187">例如，以下代码为"如果类型为'javascript'，则 `resid` 为'JSRuntime.Url'"。Outlook桌面需要此元素用于[LaunchEvent 扩展点](../../reference/manifest/extensionpoint.md#launchevent)处理程序。</span><span class="sxs-lookup"><span data-stu-id="d538e-187">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="d538e-188">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="d538e-188">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="d538e-189">语法</span><span class="sxs-lookup"><span data-stu-id="d538e-189">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="d538e-190">包含于</span><span class="sxs-lookup"><span data-stu-id="d538e-190">Contained in</span></span>

- [<span data-ttu-id="d538e-191">运行时</span><span class="sxs-lookup"><span data-stu-id="d538e-191">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="d538e-192">属性</span><span class="sxs-lookup"><span data-stu-id="d538e-192">Attributes</span></span>

|<span data-ttu-id="d538e-193">属性</span><span class="sxs-lookup"><span data-stu-id="d538e-193">Attribute</span></span>|<span data-ttu-id="d538e-194">类型</span><span class="sxs-lookup"><span data-stu-id="d538e-194">Type</span></span>|<span data-ttu-id="d538e-195">必需</span><span class="sxs-lookup"><span data-stu-id="d538e-195">Required</span></span>|<span data-ttu-id="d538e-196">说明</span><span class="sxs-lookup"><span data-stu-id="d538e-196">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d538e-197">**类型**</span><span class="sxs-lookup"><span data-stu-id="d538e-197">**type**</span></span>|<span data-ttu-id="d538e-198">string</span><span class="sxs-lookup"><span data-stu-id="d538e-198">string</span></span>|<span data-ttu-id="d538e-199">是</span><span class="sxs-lookup"><span data-stu-id="d538e-199">Yes</span></span>|<span data-ttu-id="d538e-200">指定此替代的语言。</span><span class="sxs-lookup"><span data-stu-id="d538e-200">Specifies the language for this override.</span></span> <span data-ttu-id="d538e-201">目前， `"javascript"` 是唯一受支持的选项。</span><span class="sxs-lookup"><span data-stu-id="d538e-201">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="d538e-202">**resid**</span><span class="sxs-lookup"><span data-stu-id="d538e-202">**resid**</span></span>|<span data-ttu-id="d538e-203">string</span><span class="sxs-lookup"><span data-stu-id="d538e-203">string</span></span>|<span data-ttu-id="d538e-204">是</span><span class="sxs-lookup"><span data-stu-id="d538e-204">Yes</span></span>|<span data-ttu-id="d538e-205">指定 JavaScript 文件的 URL 位置，该文件应替代在父 [Runtime](runtime.md) 元素 中定义的默认 HTML 的 URL 位置 `resid` 。</span><span class="sxs-lookup"><span data-stu-id="d538e-205">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="d538e-206">`resid`不能超过 32 个字符，并且必须与 元素中的 `id` `Url` 元素的 属性 `Resources` 匹配。</span><span class="sxs-lookup"><span data-stu-id="d538e-206">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="d538e-207">示例</span><span class="sxs-lookup"><span data-stu-id="d538e-207">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="d538e-208">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d538e-208">See also</span></span>

- [<span data-ttu-id="d538e-209">运行时</span><span class="sxs-lookup"><span data-stu-id="d538e-209">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="d538e-210">配置Outlook加载项进行基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="d538e-210">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
