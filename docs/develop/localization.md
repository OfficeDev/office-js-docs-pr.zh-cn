---
title: Office 加载项的本地化
description: 使用 Office JavaScript API 确定区域设置，并基于 Office 应用程序区域设置显示字符串，或者根据数据区域设置解释或显示数据。
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: b49d64f2c9391539ac2d5929ebff2a4ecc08b630
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349824"
---
# <a name="localization-for-office-add-ins"></a><span data-ttu-id="6c90a-103">Office 加载项的本地化</span><span class="sxs-lookup"><span data-stu-id="6c90a-103">Localization for Office Add-ins</span></span>

<span data-ttu-id="6c90a-104">您可以实现适合 Office 外接程序的任何本地化方案。</span><span class="sxs-lookup"><span data-stu-id="6c90a-104">You can implement any localization scheme that's appropriate for your Office Add-in.</span></span> <span data-ttu-id="6c90a-105">Office 外接程序平台的 JavaScript API 和清单架构提供了一些选择。</span><span class="sxs-lookup"><span data-stu-id="6c90a-105">The JavaScript API and manifest schema of the Office Add-ins platform provide some choices.</span></span> <span data-ttu-id="6c90a-106">您可以使用 Office JavaScript API 确定区域设置，并基于 Office 应用程序区域设置显示字符串，或者根据数据区域设置解释或显示数据。</span><span class="sxs-lookup"><span data-stu-id="6c90a-106">You can use the Office JavaScript API to determine a locale and display strings based on the locale of the Office application, or to interpret or display data based on the locale of the data.</span></span> <span data-ttu-id="6c90a-107">可以使用清单指定区域设置特定的加载项文件位置和描述性信息。</span><span class="sxs-lookup"><span data-stu-id="6c90a-107">You can use the manifest to specify locale-specific add-in file location and descriptive information.</span></span> <span data-ttu-id="6c90a-108">也可以使用 Microsoft Ajax 脚本支持全球化和本地化。</span><span class="sxs-lookup"><span data-stu-id="6c90a-108">Alternatively, you can use Microsoft Ajax script to support globalization and localization.</span></span>

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a><span data-ttu-id="6c90a-109">使用 JavaScript API 确定区域设置特定的字符串</span><span class="sxs-lookup"><span data-stu-id="6c90a-109">Use the JavaScript API to determine locale-specific strings</span></span>

<span data-ttu-id="6c90a-110">JavaScript API Office两个属性支持显示或解释与应用程序和数据区域设置一Office值：</span><span class="sxs-lookup"><span data-stu-id="6c90a-110">The Office JavaScript API provides two properties that support displaying or interpreting values consistent with the locale of the Office application and data:</span></span>

- <span data-ttu-id="6c90a-111">[Context.displayLanguage][displayLanguage]指定 (应用程序) 用户界面区域设置Office语言。</span><span class="sxs-lookup"><span data-stu-id="6c90a-111">[Context.displayLanguage][displayLanguage] specifies the locale (or language) of the user interface of the Office application.</span></span> <span data-ttu-id="6c90a-112">以下示例验证应用程序是否Office en-US 或 fr-FR 区域设置，并显示区域设置特定的问候语。</span><span class="sxs-lookup"><span data-stu-id="6c90a-112">The following example verifies if the Office application uses the en-US or fr-FR locale, and displays a locale-specific greeting.</span></span>

    ```js
    function sayHelloWithDisplayLanguage() {
        var myLanguage = Office.context.displayLanguage;
        switch (myLanguage) {
            case 'en-US':
                write('Hello!');
                break;
            case 'fr-FR':
                write('Bonjour!');
                break;
        }
    }

    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }
    ```

- <span data-ttu-id="6c90a-p103">[Context.contentLanguage][contentLanguage] 指定数据的区域设置（或语言）。展开上一个代码示例，不检查 [displayLanguage] 属性，而是为 `myLanguage` 分配 [contentLanguage] 属性值，并使用相同代码的其余部分根据数据的区域设置显示问候语：</span><span class="sxs-lookup"><span data-stu-id="6c90a-p103">[Context.contentLanguage][contentLanguage] specifies the locale (or language) of the data. Extending the last code sample, instead of checking the [displayLanguage] property, assign `myLanguage` the value of the [contentLanguage] property, and use the rest of the same code to display a greeting based on the locale of the data:</span></span>

    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a><span data-ttu-id="6c90a-115">通过清单控制本地化</span><span class="sxs-lookup"><span data-stu-id="6c90a-115">Control localization from the manifest</span></span>


<span data-ttu-id="6c90a-116">每个 Office 外接程序在其清单中指定一个 [DefaultLocale] 元素和区域设置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-116">Every Office Add-in specifies a [DefaultLocale] element and a locale in its manifest.</span></span> <span data-ttu-id="6c90a-117">默认情况下，Office外接程序平台和 Office 客户端应用程序将[]Description、DisplayName、IconUrl、HighResolutionIconUrl 和[] [SourceLocation]元素的值应用于所有区域设置。 [] []</span><span class="sxs-lookup"><span data-stu-id="6c90a-117">By default, the Office Add-in platform and Office client applications apply the values of the [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl], and [SourceLocation] elements to all locales.</span></span> <span data-ttu-id="6c90a-118">可以通过为每个其他区域设置的上述五个元素中的任意一个指定 [Override] 子元素来选择支持将特定值用于特定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-118">You can optionally support specific values for specific locales, by specifying an [Override] child element for each additional locale, for any of these five elements.</span></span> <span data-ttu-id="6c90a-119">[DefaultLocale] 元素和 `Locale` 元素的 [] 属性的值根据 [RFC 3066]（“用于语言标识的标记”）指定。</span><span class="sxs-lookup"><span data-stu-id="6c90a-119">The value for the [DefaultLocale] element and for the `Locale` attribute of the [Override] element is specified according to [RFC 3066], "Tags for the Identification of Languages."</span></span> <span data-ttu-id="6c90a-120">表 1 描述了这些元素的本地化支持。</span><span class="sxs-lookup"><span data-stu-id="6c90a-120">Table 1 describes the localizing support for these elements.</span></span>

<span data-ttu-id="6c90a-121">*表 1.本地化支持*</span><span class="sxs-lookup"><span data-stu-id="6c90a-121">*Table 1. Localization support*</span></span>


|<span data-ttu-id="6c90a-122">**Element**</span><span class="sxs-lookup"><span data-stu-id="6c90a-122">**Element**</span></span>|<span data-ttu-id="6c90a-123">**本地化支持**</span><span class="sxs-lookup"><span data-stu-id="6c90a-123">**Localization support**</span></span>|
|:-----|:-----|
|<span data-ttu-id="6c90a-124">[Description]</span><span class="sxs-lookup"><span data-stu-id="6c90a-124">[Description]</span></span>   |<span data-ttu-id="6c90a-125">指定的每个区域设置中的用户都可以在 AppSource（或专有目录）中看到本地化的加载项说明。</span><span class="sxs-lookup"><span data-stu-id="6c90a-125">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="6c90a-126">对于 Outlook 加载项，在安装后，用户可以在 Exchange 管理中心 (EAC) 中看到说明。</span><span class="sxs-lookup"><span data-stu-id="6c90a-126">For Outlook add-ins, users can see the description in the Exchange Admin Center (EAC) after installation.</span></span>|
|<span data-ttu-id="6c90a-127">[DisplayName]</span><span class="sxs-lookup"><span data-stu-id="6c90a-127">[DisplayName]</span></span>   |<span data-ttu-id="6c90a-128">指定的每个区域设置中的用户都可以在 AppSource（或专有目录）中看到本地化的加载项说明。</span><span class="sxs-lookup"><span data-stu-id="6c90a-128">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="6c90a-129">对于 Outlook 加载项，在安装后，用户可以看到显示名称为 Outlook 加载项按钮标签，也可以在 EAC 中看到显示名称。</span><span class="sxs-lookup"><span data-stu-id="6c90a-129">For Outlook add-ins, users can see the display name as a label for the Outlook add-in button and in the EAC after installation.</span></span><br/><span data-ttu-id="6c90a-130">对于内容和任务窗格外接程序，安装外接程序后，用户可以在功能区中看到该显示名称。</span><span class="sxs-lookup"><span data-stu-id="6c90a-130">For content and task pane add-ins, users can see the display name in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="6c90a-131">[IconUrl]</span><span class="sxs-lookup"><span data-stu-id="6c90a-131">[IconUrl]</span></span>        |<span data-ttu-id="6c90a-p105">图标图像是可选的。可以使用相同的替代方法为特定区域性指定特定图像。如果使用并本地化图标，则您指定的每个区域设置中的用户均可看到该加载项的本地化图标图像。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p105">The icon image is optional. You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="6c90a-135">对于 Outlook 外接程序，安装外接程序后，用户可以在 EAC 中看到该图标。</span><span class="sxs-lookup"><span data-stu-id="6c90a-135">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="6c90a-136">对于内容和任务窗格加载项，安装加载项后，用户可以在功能区中看到此图标。</span><span class="sxs-lookup"><span data-stu-id="6c90a-136">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="6c90a-137">[HighResolutionIconUrl] **重要说明：** 此元素仅适用于使用加载项清单版本 1.1 的情况。</span><span class="sxs-lookup"><span data-stu-id="6c90a-137">[HighResolutionIconUrl] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>|<span data-ttu-id="6c90a-p106">高分辨率图标图像是可选的，但一旦指定，则必须在  [IconUrl] 元素之后出现。指定 [HighResolutionIconUrl] 且在支持高 DPI 分辨率的设备上安装了加载项后，将使用 [HighResolutionIconUrl] 值而不是 [IconUrl] 值。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p106">The high resolution icon image is optional but if it is specified, it must occur after the  [IconUrl] element. When [HighResolutionIconUrl] is specified, and the add-in is installed on a device that supports high dpi resolution, the [HighResolutionIconUrl] value is used instead of the value for [IconUrl].</span></span><br/><span data-ttu-id="6c90a-p107">图标图像是可选的。可以使用相同的替代方法为特定区域性指定特定图像。如果使用并本地化图标，则您指定的每个区域设置中的用户均可看到该加载项的本地化图标图像。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p107">You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="6c90a-142">对于 Outlook 外接程序，安装外接程序后，用户可以在 EAC 中看到该图标。</span><span class="sxs-lookup"><span data-stu-id="6c90a-142">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="6c90a-143">对于内容和任务窗格加载项，安装加载项后，用户可以在功能区中看到此图标。</span><span class="sxs-lookup"><span data-stu-id="6c90a-143">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="6c90a-144">[Resources] **重要说明：** 此元素仅适用于使用加载项清单版本 1.1 的情况。</span><span class="sxs-lookup"><span data-stu-id="6c90a-144">[Resources] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>   |<span data-ttu-id="6c90a-145">指定的每个区域设置中的用户都可以看到专门针对相应区域设置为加载项创建的字符串和图标资源。</span><span class="sxs-lookup"><span data-stu-id="6c90a-145">Users in each locale you specify can see string and icon resources that you specifically create for the add-in for that locale.</span></span> |
|<span data-ttu-id="6c90a-146">[SourceLocation]</span><span class="sxs-lookup"><span data-stu-id="6c90a-146">[SourceLocation]</span></span>   |<span data-ttu-id="6c90a-147">指定的每个区域设置中的用户都可以看到专门针对该区域设置为该加载项设计的网页。</span><span class="sxs-lookup"><span data-stu-id="6c90a-147">Users in each locale you specify can see a webpage that you specifically design for the add-in for that locale.</span></span> |


> [!NOTE]
> <span data-ttu-id="6c90a-148">可以仅为 Office 支持的语言环境对说明和显示名称进行本地化。</span><span class="sxs-lookup"><span data-stu-id="6c90a-148">You can localize the description and display name for only the locales that Office supports.</span></span> <span data-ttu-id="6c90a-149">有关当前版本的 Office 的语言和区域设置列表，请参阅 [Office 2013 中的语言标识符和 OptionState ID 值](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))。</span><span class="sxs-lookup"><span data-stu-id="6c90a-149">See [Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) for a list of languages and locales for the current release of Office.</span></span>


### <a name="examples"></a><span data-ttu-id="6c90a-150">示例</span><span class="sxs-lookup"><span data-stu-id="6c90a-150">Examples</span></span>

<span data-ttu-id="6c90a-p109">例如，Office 加载项可以将 [DefaultLocale] 指定为 `en-us`。对于 [DisplayName] 元素，加载项可以为区域设置 `fr-fr` 指定 [Override] 子元素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p109">For example, an Office Add-in can specify the [DefaultLocale] as `en-us`. For the [DisplayName] element, the add-in can specify an [Override] child element for the locale `fr-fr`, as shown below.</span></span>


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> <span data-ttu-id="6c90a-153">如需针对一个语系内的多个区域进行本地化，例如 `de-de` 和 `de-at`，则建议对各个区域使用独立的 `Override` 元素。</span><span class="sxs-lookup"><span data-stu-id="6c90a-153">If you need to localize for more than one area within a language family, such as `de-de` and `de-at`, we recommend that you use separate `Override` elements for each area.</span></span> <span data-ttu-id="6c90a-154">在客户端应用程序和平台的所有组合中，不支持仅单独使用语言 `de` Office，在这种情况下， 。</span><span class="sxs-lookup"><span data-stu-id="6c90a-154">Using just the language name alone, in this case, `de`, is not supported across all combinations of Office client applications and platforms.</span></span>

<span data-ttu-id="6c90a-p111">这意味着，加载项默认情况下采用 `en-us` 区域设置。除非客户端计算机的区域设置为 `fr-fr`（此时用户将看到法语的显示名称“Lecteur vidéo”），否则对于所有区域设置，用户都将看到英文显示名称“Video player”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p111">This means that the add-in assumes the  `en-us` locale by default. Users see the English display name of "Video player" for all locales unless the client computer's locale is `fr-fr`, in which case users would see the French display name "Lecteur vidéo".</span></span>

> [!NOTE]
> <span data-ttu-id="6c90a-157">每种语言只可指定单一的覆盖，包括对于默认区域设置的覆盖。</span><span class="sxs-lookup"><span data-stu-id="6c90a-157">You may only specify a single override per language, including for the default locale.</span></span> <span data-ttu-id="6c90a-158">例如，如果默认区域设置为 `en-us`，则无法也指定 `en-us` 的覆盖。</span><span class="sxs-lookup"><span data-stu-id="6c90a-158">For example, if your default locale is `en-us` you cannot not specify an  override for `en-us` as well.</span></span> 

<span data-ttu-id="6c90a-p113">以下示例对 [Description] 元素应用区域设置覆盖。它首先指定默认区域设置 `en-us` 和英文说明，然后指定 [Override] 语句，其中包含 `fr-fr` 区域设置的法语说明：</span><span class="sxs-lookup"><span data-stu-id="6c90a-p113">The following example applies a locale override for the [Description] element. It first specifies a default locale of `en-us` and an English description, and then specifies an [Override] statement with a French description for the `fr-fr` locale:</span></span>

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook."/>
</Description>
```

<span data-ttu-id="6c90a-p114">也就是说，加载项默认采用 `en-us` 区域设置。除非客户端计算机的区域设置为 `fr-fr`（此时用户将看到法语说明），否则对于所有区域设置，用户都将看到 `DefaultValue` 属性中的英文说明。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p114">This means that the add-in assumes the `en-us` locale by default. Users would see the English description in the `DefaultValue` attribute for all locales unless the client computer's locale is `fr-fr`, in which case they would see the French description.</span></span>

<span data-ttu-id="6c90a-p115">在以下示例中，加载项指定更适合 `fr-fr` 区域设置和区域性的不同图像。默认情况下，用户看到的是图像 DefaultLogo.png，客户端计算机的区域设置为 `fr-fr` 时除外。此时，用户将看到图像 FrenchLogo.png。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p115">In the following example, the add-in specifies a separate image that's more appropriate for the `fr-fr` locale and culture. Users see the image DefaultLogo.png by default, except when the locale of the client computer is `fr-fr`. In this case, users would see the image FrenchLogo.png.</span></span> 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

<span data-ttu-id="6c90a-p116">以下示例显示了如何本地化 `Resources` 部分中的资源。它对一个更适用于 `ja-jp` 區域性的图像应用了区域设置覆盖。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p116">The following example shows how to localize a resource in the `Resources` section. It applies a locale override for an image that is more appropriate for the `ja-jp` culture.</span></span>

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


<span data-ttu-id="6c90a-p117">对于 [SourceLocation] 元素，支持其他区域设置意味着为每个指定的区域设置提供单独的源 HTML 文件。指定的每个区域设置中的用户都可以看到为相应区域设置设计的自定义网页。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p117">For the [SourceLocation] element, supporting additional locales means providing a separate source HTML file for each of the specified locales. Users in each locale you specify can see a customized webpage that you design for that them.</span></span>

<span data-ttu-id="6c90a-p118">对于 Outlook 加载项，[SourceLocation] 元素还与外形规格保持一致。这样一来，就可以为每个相应外形规格提供不同的本地化源 HTML 文件。可以在每个适用的 settings 元素（[DesktopSettings]、[TabletSettings] 或 [PhoneSettings]）中指定一个或多个 [Override] 子元素。下面的示例展示了用于台式机、平板电脑和智能手机外形规格的 settings 元素，每个都有一个用于默认区域设置的 HTML 文件，以及另一个用于法语区域设置的 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p118">For Outlook add-ins, the [SourceLocation] element also aligns to the form factor. This allows you to provide a separate, localized source HTML file for each corresponding form factor. You can specify one or more [Override] child elements in each applicable settings element ([DesktopSettings], [TabletSettings], or [PhoneSettings]). The following example shows settings elements for the desktop, tablet, and smartphone form factors, each with one HTML file for the default locale and another for the French locale.</span></span>


```xml
<DesktopSettings>
   <SourceLocation DefaultValue="https://contoso.com/Desktop.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Desktop.html" />
   </SourceLocation>
   <RequestedHeight>250</RequestedHeight>
</DesktopSettings>
<TabletSettings>
   <SourceLocation DefaultValue="https://contoso.com/Tablet.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Tablet.html" />
   </SourceLocation>
   <RequestedHeight>200</RequestedHeight>
</TabletSettings>
<PhoneSettings>
   <SourceLocation DefaultValue="https://contoso.com/Mobile.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Mobile.html" />
   </SourceLocation>
</PhoneSettings>
```

## <a name="localize-extended-overrides"></a><span data-ttu-id="6c90a-174">本地化扩展替代</span><span class="sxs-lookup"><span data-stu-id="6c90a-174">Localize extended overrides</span></span>

<span data-ttu-id="6c90a-175">Office 外接程序的一些扩展性功能（如键盘快捷方式）使用托管在服务器上的 JSON 文件进行配置，而不是使用外接程序的 XML 清单进行配置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-175">Some extensibility features of Office Add-ins, such as keyboard shortcuts, are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span> <span data-ttu-id="6c90a-176">本节假定你熟悉扩展替代。</span><span class="sxs-lookup"><span data-stu-id="6c90a-176">This section assumes that you're familiar with extended overrides.</span></span> <span data-ttu-id="6c90a-177">请参阅 [使用清单和](extended-overrides.md) [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素的扩展替代。</span><span class="sxs-lookup"><span data-stu-id="6c90a-177">See [Work with extended overrides of the manifest](extended-overrides.md) and [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span>

<span data-ttu-id="6c90a-178">使用 `ResourceUrl` [ExtendedOverrides](../reference/manifest/extendedoverrides.md)元素的 属性Office本地化资源的文件。</span><span class="sxs-lookup"><span data-stu-id="6c90a-178">Use the `ResourceUrl` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element to point Office to a file of localized resources.</span></span> <span data-ttu-id="6c90a-179">示例如下。</span><span class="sxs-lookup"><span data-stu-id="6c90a-179">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="6c90a-180">然后，扩展替代文件使用令牌而不是字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-180">The extended overrides file then uses tokens instead of strings.</span></span> <span data-ttu-id="6c90a-181">令牌命名资源文件中字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-181">The tokens name strings in the resource file.</span></span> <span data-ttu-id="6c90a-182">下面的示例将键盘快捷方式分配给 (加载项任务窗格) 定义的函数。</span><span class="sxs-lookup"><span data-stu-id="6c90a-182">The following is an example that assigns a keyboard shortcut to a function (defined elsewhere) that displays the add-in's task pane.</span></span> <span data-ttu-id="6c90a-183">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="6c90a-183">Note about this markup:</span></span>

- <span data-ttu-id="6c90a-184">该示例并不十分有效。</span><span class="sxs-lookup"><span data-stu-id="6c90a-184">The example isn't quite valid.</span></span> <span data-ttu-id="6c90a-185"> (我们在下面向它添加所需的附加属性。) </span><span class="sxs-lookup"><span data-stu-id="6c90a-185">(We add a required additional property to it below.)</span></span>
- <span data-ttu-id="6c90a-186">令牌的格式必须为 **${resource。*name-of-resource*}**.</span><span class="sxs-lookup"><span data-stu-id="6c90a-186">The tokens must have the format **${resource.*name-of-resource*}**.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ] 
}
```

<span data-ttu-id="6c90a-187">资源文件也是 JSON 格式的文件，其顶级属性按区域设置划分为子 `resources` 属性。</span><span class="sxs-lookup"><span data-stu-id="6c90a-187">The resource file, which is also JSON-formatted, has a top-level `resources` property that is divided into subproperties by locale.</span></span> <span data-ttu-id="6c90a-188">对于每个区域设置，为扩展替代文件中使用的每个令牌分配一个字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-188">For each locale, a string is assigned to each token that was used in the extended overrides file.</span></span> <span data-ttu-id="6c90a-189">下面是一个包含 和 的字符串 `en-us` 的示例 `fr-fr` 。</span><span class="sxs-lookup"><span data-stu-id="6c90a-189">The following is an example which has strings for `en-us` and `fr-fr`.</span></span> <span data-ttu-id="6c90a-190">本示例中，键盘快捷方式在两个区域设置中都相同，但并非总是如此，尤其是在本地化具有不同字母或书写系统，因而具有不同键盘区域设置时。</span><span class="sxs-lookup"><span data-stu-id="6c90a-190">In this example, the keyboard shortcut is the same in both locales, but that won't always be the case, especially when you are localizing for locales that have a different alphabet or writing system, and hence a different keyboard.</span></span>

```json
{
    "resources":{ 
        "en-us": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            }, 
        },
        "fr-fr": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Afficher le volet de tâche pour add-in",
              } 
        }
    }
}
```

<span data-ttu-id="6c90a-191">文件中没有与 和 节 `default` 对等 `en-us` `fr-fr` 的属性。</span><span class="sxs-lookup"><span data-stu-id="6c90a-191">There is no `default` property in the file that is a peer to the `en-us` and `fr-fr` sections.</span></span> <span data-ttu-id="6c90a-192">这是因为，当 Office 主机应用程序区域设置与资源文件的任何 *ll-cc* 属性不匹配时，必须使用扩展替代文件本身定义默认 *字符串*。</span><span class="sxs-lookup"><span data-stu-id="6c90a-192">This is because the default strings, which are used when the locale of the Office host application doesn't match any of the *ll-cc* properties in the resources file, *must be defined in the extended overrides file itself*.</span></span> <span data-ttu-id="6c90a-193">直接在扩展替代文件中定义默认字符串可确保当 Office 应用程序区域设置与清单) 中指定的外接程序 (的默认区域设置相匹配时，Office 不会下载资源文件。</span><span class="sxs-lookup"><span data-stu-id="6c90a-193">Defining the default strings directly in the extended overrides file ensures that Office doesn't download the resource file when the locale of the Office application matches the default locale of the add-in (as specified in the manifest).</span></span> <span data-ttu-id="6c90a-194">下面是前面使用资源令牌的扩展替代文件的更正版本。</span><span class="sxs-lookup"><span data-stu-id="6c90a-194">The following is a corrected version of the preceding example of an extended overrides file that uses resource tokens.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ],
    "resources": { 
        "default": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            } 
        }
    }
}
```

## <a name="match-datetime-format-with-client-locale"></a><span data-ttu-id="6c90a-195">将日期/时间格式与客户端区域设置匹配</span><span class="sxs-lookup"><span data-stu-id="6c90a-195">Match date/time format with client locale</span></span>

<span data-ttu-id="6c90a-196">可以使用 **[displayLanguage]** 属性获取 Office 应用程序用户界面区域设置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-196">You can get the locale of the user interface of the Office client application by using the **[displayLanguage]** property.</span></span> <span data-ttu-id="6c90a-197">然后，您可以按照与应用程序当前区域设置一致的格式显示日期和时间Office值。</span><span class="sxs-lookup"><span data-stu-id="6c90a-197">You can then display date and time values in a format consistent with the current locale of the Office application.</span></span> <span data-ttu-id="6c90a-198">执行上述操作的一种方法是准备一个指定日期/时间显示格式的资源文件以用于 Office 外界程序支持的各个区域设置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-198">One way to do that is to prepare a resource file that specifies the date/time display format to use for each locale that your Office Add-in supports.</span></span> <span data-ttu-id="6c90a-199">运行时，加载项可以使用资源文件，将相应的日期/时间格式与从 **[displayLanguage]** 属性获取区域设置相匹配。</span><span class="sxs-lookup"><span data-stu-id="6c90a-199">At run time, your add-in can use the resource file and match the appropriate date/time format with the locale obtained from the **[displayLanguage]** property.</span></span>

<span data-ttu-id="6c90a-200">可以使用[contentLanguage]属性获取 Office 应用程序的数据区域设置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-200">You can get the locale of the data of the Office client application by using the [contentLanguage] property.</span></span> <span data-ttu-id="6c90a-201">基于此值，可以正确地解读或显示日期/时间字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-201">Based on this value, you can then appropriately interpret or display date/time strings.</span></span> <span data-ttu-id="6c90a-202">例如，`jp-JP` 区域设置将数据/时间值表示为 `yyyy/MM/dd`，而 `fr-FR` 区域设置则表示为 `dd/MM/yyyy`。</span><span class="sxs-lookup"><span data-stu-id="6c90a-202">For example, the `jp-JP` locale expresses data/time values as `yyyy/MM/dd`, and the `fr-FR` locale, `dd/MM/yyyy`.</span></span>


## <a name="use-ajax-for-globalization-and-localization"></a><span data-ttu-id="6c90a-203">将 Ajax 用于全球化和本地化</span><span class="sxs-lookup"><span data-stu-id="6c90a-203">Use Ajax for globalization and localization</span></span>


<span data-ttu-id="6c90a-204">如果使用 Visual Studio 创建 Office 外接程序，.NET Framework 和 Ajax 会提供用于全球化和本地化客户端脚本文件的方法。</span><span class="sxs-lookup"><span data-stu-id="6c90a-204">If you use Visual Studio to create Office Add-ins, the .NET Framework and Ajax provide ways to globalize and localize client script files.</span></span>

<span data-ttu-id="6c90a-p127">您可以全球化 Office 外接程序并在其 JavaScript 代码中使用 [Date](/previous-versions/bb310850(v=vs.140)) 和 [Number](/previous-versions/bb310835(v=vs.140)) JavaScript 类型扩展和 JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) 对象，以根据当前浏览器的区域设置显示值。有关详细信息，请参阅 [Walkthrough: Globalizing a Date by Using Client Script](/previous-versions/bb386581(v=vs.140))。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p127">You can globalize and use the [Date](/previous-versions/bb310850(v=vs.140)) and [Number](/previous-versions/bb310835(v=vs.140)) JavaScript type extensions and the JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object in the JavaScript code for an Office Add-in to display values based on the locale settings on the current browser. For more information, see [Walkthrough: Globalizing a Date by Using Client Script](/previous-versions/bb386581(v=vs.140)).</span></span>

<span data-ttu-id="6c90a-p128">可将本地化的资源字符串直接包含在独立的 JavaScript 文件中，以便为不同区域设置提供客户端脚本文件，这些文件在浏览器中设置或由用户提供。为每个受支持的区域设置创建单独的脚本文件。在每个脚本文件中，包括一个 JSON 格式的对象，其中包含用于该区域设置的资源字符串。在浏览器中运行脚本时，会应用本地化的值。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p128">You can include localized resource strings directly in standalone JavaScript files to provide client script files for different locales, which are set on the browser or provided by the user. Create a separate script file for each supported locale. In each script file, include an object in JSON format that contains the resource strings for that locale. The localized values are applied when the script runs in the browser.</span></span>


## <a name="example-build-a-localized-office-add-in"></a><span data-ttu-id="6c90a-211">示例：生成本地化 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6c90a-211">Example: Build a localized Office Add-in</span></span>

<span data-ttu-id="6c90a-212">本节提供示例，演示如何本地化 Office 外接程序描述、显示名称和 UI。</span><span class="sxs-lookup"><span data-stu-id="6c90a-212">This section provides examples that show you how to localize an Office Add-in description, display name, and UI.</span></span> 

> [!NOTE]
> <span data-ttu-id="6c90a-213">若要下载 Visual Studio 2019，请参阅 Visual Studio [IDE 页面](https://visualstudio.microsoft.com/vs/)。</span><span class="sxs-lookup"><span data-stu-id="6c90a-213">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="6c90a-214">在安装过程中，你需要选择 Office/SharePoint 开发工作负载。</span><span class="sxs-lookup"><span data-stu-id="6c90a-214">During installation you'll need to select the Office/SharePoint development workload.</span></span>

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a><span data-ttu-id="6c90a-215">配置 Office 以使用其他语言进行显示或编辑</span><span class="sxs-lookup"><span data-stu-id="6c90a-215">Configure Office to use additional languages for display or editing</span></span>

<span data-ttu-id="6c90a-216">若要运行提供的示例代码，请在计算机上将 Office 配置为使用其他语言，以便可以通过切换用于菜单和命令显示的语言、或切换用于编辑和校对的语言或同时切换两者来测试外接程序。</span><span class="sxs-lookup"><span data-stu-id="6c90a-216">To run the sample code provided, configure Office on your computer to use additional languages so that you can test your add-in by switching the language used for display in menus and commands, for editing and proofing, or both.</span></span>

<span data-ttu-id="6c90a-217">可以使用 Office 语言包安装其他语言。</span><span class="sxs-lookup"><span data-stu-id="6c90a-217">You can use an Office Language pack to install an additional language.</span></span> <span data-ttu-id="6c90a-218">有关语言包以及如何获取语言包的详细信息，请参阅[适用于 Office 的 Language Accessory Pack](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f)。</span><span class="sxs-lookup"><span data-stu-id="6c90a-218">For more information about Language Packs and where to get them, see [Language Accessory Pack for Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).</span></span>

<span data-ttu-id="6c90a-219">安装 Language Accessory Pack 后，可以将 Office 配置使用安装的语言，以便在 UI 中显示，或编辑文档内容，或者两者兼具。</span><span class="sxs-lookup"><span data-stu-id="6c90a-219">After you install the Language Accessory Pack, you can configure Office to use the installed language for display in the UI, for editing document content, or both.</span></span> <span data-ttu-id="6c90a-220">在本文的示例中，所安装的 Office 应用了西班牙语语言包。</span><span class="sxs-lookup"><span data-stu-id="6c90a-220">The example in this article uses an installation of Office that has the Spanish Language Pack applied.</span></span>

### <a name="create-an-office-add-in-project"></a><span data-ttu-id="6c90a-221">创建 Office 加载项项目</span><span class="sxs-lookup"><span data-stu-id="6c90a-221">Create an Office Add-in project</span></span>

<span data-ttu-id="6c90a-222">您需要创建一个 2019 Visual Studio 2019 Office外接程序项目。</span><span class="sxs-lookup"><span data-stu-id="6c90a-222">You'll need to create a Visual Studio 2019 Office Add-in project.</span></span>

> [!NOTE]
> <span data-ttu-id="6c90a-223">如果您尚未安装 Visual Studio 2019，请参阅[Visual Studio IDE 页面，了解](https://visualstudio.microsoft.com/vs/)下载说明。</span><span class="sxs-lookup"><span data-stu-id="6c90a-223">If you haven't installed Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/) for download instructions.</span></span> <span data-ttu-id="6c90a-224">在安装过程中，你需要选择 Office/SharePoint 开发工作负载。</span><span class="sxs-lookup"><span data-stu-id="6c90a-224">During installation you'll need to select the Office/SharePoint development workload.</span></span> <span data-ttu-id="6c90a-225">如果之前已安装 Visual Studio 2019，Visual Studio 安装程序确保安装了[](/visualstudio/install/modify-visual-studio/)Office/SharePoint 开发工作负载。</span><span class="sxs-lookup"><span data-stu-id="6c90a-225">If you have previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio/) to ensure that the Office/SharePoint development workload is installed.</span></span>

1. <span data-ttu-id="6c90a-226">选择“**创建新项目**”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-226">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="6c90a-227">使用搜索框，输入“**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-227">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="6c90a-228">选择“**Word Web 外接程序**”，然后选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-228">Choose **Word Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="6c90a-229">将项目命名 **为 WorldReadyAddIn，** 然后选择"**创建"。**</span><span class="sxs-lookup"><span data-stu-id="6c90a-229">Name your project **WorldReadyAddIn** and select **Create**.</span></span>

4. <span data-ttu-id="6c90a-p134">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p134">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>


### <a name="localize-the-text-used-in-your-add-in"></a><span data-ttu-id="6c90a-232">本地化加载项中使用的文本</span><span class="sxs-lookup"><span data-stu-id="6c90a-232">Localize the text used in your add-in</span></span>

<span data-ttu-id="6c90a-233">您要本地化为另一种语言的文本出现在两个区域中：</span><span class="sxs-lookup"><span data-stu-id="6c90a-233">The text that you want to localize for another language appears in two areas:</span></span>

-  <span data-ttu-id="6c90a-p135">**加载项显示名称和说明**。这是受应用程序清单文件中的条目控制的。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p135">**Add-in display name and description**. This is controlled by entries in the add-in manifest file.</span></span>

-  <span data-ttu-id="6c90a-236">**加载项 UI**。</span><span class="sxs-lookup"><span data-stu-id="6c90a-236">**Add-in UI**.</span></span> <span data-ttu-id="6c90a-237">可以通过使用 JavaScript 代码本地化在加载项 UI 中出现的字符串（例如，通过使用包含已本地化的字符串的单独资源文件）。</span><span class="sxs-lookup"><span data-stu-id="6c90a-237">You can localize the strings that appear in your add-in UI by using JavaScript code, for example, by using a separate resource file that contains the localized strings.</span></span>

<span data-ttu-id="6c90a-238">要本地化加载项显示名称和说明，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6c90a-238">To localize the add-in display name and description:</span></span>

1. <span data-ttu-id="6c90a-239">在“解决方案资源管理器”中，展开“WorldReadyAddIn”、“WorldReadyAddInManifest”，然后选择“WorldReadyAddIn.xml”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-239">In **Solution Explorer**, expand **WorldReadyAddIn**, **WorldReadyAddInManifest**, and then choose **WorldReadyAddIn.xml**.</span></span>

2. <span data-ttu-id="6c90a-240">在WorldReadyAddInManifest.xml中，将 [DisplayName] 和 [Description] 元素替换为以下代码块。</span><span class="sxs-lookup"><span data-stu-id="6c90a-240">In WorldReadyAddInManifest.xml, replace the [DisplayName] and [Description] elements with the following block of code.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6c90a-241">对于本示例中使用的西班牙语本地化字符串的[DisplayName] 和 [Description] 元素，您可以替换为任何其他语言的本地化字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-241">You can replace the Spanish language localized strings used in this example for the [DisplayName] and [Description] elements with the localized strings for any other language.</span></span>

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. <span data-ttu-id="6c90a-242">例如，如果您将 Office 2013 的显示语言从英语切换到西班牙语，然后运行加载项，加载项的显示名称和说明将用本地化文本显示。</span><span class="sxs-lookup"><span data-stu-id="6c90a-242">When you change the display language for Office 2013 from English to Spanish, for example, and then run the add-in, the add-in display name and description are shown with localized text.</span></span>

<span data-ttu-id="6c90a-243">若要设计加载项 UI 的布局，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6c90a-243">To lay out the add-in UI:</span></span>

1. <span data-ttu-id="6c90a-244">在 Visual Studio 的“解决方案资源管理器”中，选择“Home.html”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-244">In Visual Studio, in **Solution Explorer**, choose **Home.html**.</span></span>

2. <span data-ttu-id="6c90a-245">在 Home.html 中，将 `<body>` 元素替换为以下 HTML，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="6c90a-245">Replace the `<body>` element contents in Home.html with the following HTML, and save the file.</span></span>

    ```html
    <body>
        <!-- Page content -->
        <div id="content-header" class="ms-bgColor-themePrimary ms-font-xl">
            <div class="padding">
                <h1 id="greeting" class="ms-fontColor-white"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div class="ms-font-m">
                    <p id="about"></p>
                </div>
            </div>
        </div>
    </body>
    ```

<span data-ttu-id="6c90a-246">下图展示了在完成剩余步骤和运行加载项时显示本地化文本的 heading (h1) 元素和 paragraph (p) 元素。</span><span class="sxs-lookup"><span data-stu-id="6c90a-246">The following figure shows the heading (h1) element and the paragraph (p) element that will display localized text when you complete the remaining steps and run the add-in.</span></span>

<span data-ttu-id="6c90a-247">*图 1：加载项 UI*</span><span class="sxs-lookup"><span data-stu-id="6c90a-247">*Figure 1. The add-in UI*</span></span>

![具有突出显示部分的应用程序用户界面。](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a><span data-ttu-id="6c90a-249">添加包含本地化后字符串的资源文件</span><span class="sxs-lookup"><span data-stu-id="6c90a-249">Add the resource file that contains the localized strings</span></span>

<span data-ttu-id="6c90a-250">JavaScript 资源文件包含加载项 UI 使用的字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-250">The JavaScript resource file contains the strings used for the add-in UI.</span></span> <span data-ttu-id="6c90a-251">示例加载项 UI 的 HTML 中包含用于显示问候语的 `<h1>` 元素以及用于向用户介绍加载项的 `<p>` 元素。</span><span class="sxs-lookup"><span data-stu-id="6c90a-251">The HTML for the sample add-in UI contains an `<h1>` element that displays a greeting, and a `<p>` element that introduces the add-in to the user.</span></span> 

<span data-ttu-id="6c90a-p138">若要为标题和段落启用本地化字符串，您需要将字符串放在一个单独的资源文件中。资源文件会创建一个 JavaScript 对象，对每组本地化字符串来说，它都包含一个单独的 JavaScript 对象表示法 (JSON) 对象。资源文件也提供为给定区域设置找回适当 JSON 对象的方法。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p138">To enable localized strings for the heading and paragraph, you place the strings in a separate resource file. The resource file creates a JavaScript object that contains a separate JavaScript Object Notation (JSON) object for each set of localized strings. The resource file also provides a method for getting back the appropriate JSON object for a given locale.</span></span>

<span data-ttu-id="6c90a-255">若要为加载项项目添加资源文件：</span><span class="sxs-lookup"><span data-stu-id="6c90a-255">To add the resource file to the add-in project:</span></span>

1. <span data-ttu-id="6c90a-256">在 Visual Studio 的“解决方案资源管理器”中，右键单击“WorldReadyAddInWeb”项目并选择“添加” > “新项目”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-256">In **Solution Explorer** in Visual Studio, right-click the **WorldReadyAddInWeb** project and choose **Add** > **New Item**.</span></span> 

2. <span data-ttu-id="6c90a-257">在“添加新项目”对话框中，选择“JavaScript 文件”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-257">In the **Add New Item** dialog box, choose **JavaScript File**.</span></span>

3. <span data-ttu-id="6c90a-258">输入文件名 **UIStrings.js**，然后选择“添加”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-258">Enter **UIStrings.js** as the file name and choose **Add**.</span></span>

4. <span data-ttu-id="6c90a-259">将以下代码添加到 UIStrings.js 文件，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="6c90a-259">Add the following code to the UIStrings.js file, and save the file.</span></span>

    ```js
    /* Store the locale-specific strings */

    var UIStrings = (function ()
    {
        "use strict";

        var UIStrings = {};

        // JSON object for English strings
        UIStrings.EN =
        {
            "Greeting": "Welcome",
            "Introduction": "This is my localized add-in."
        };

        // JSON object for Spanish strings
        UIStrings.ES =
        {
            "Greeting": "Bienvenido",
            "Introduction": "Esta es mi aplicación localizada."
        };

        UIStrings.getLocaleStrings = function (locale)
        {
            var text;

            // Get the resource strings that match the language.
            switch (locale)
            {
                case 'en-US':
                    text = UIStrings.EN;
                    break;
                case 'es-ES':
                    text = UIStrings.ES;
                    break;
                default:
                    text = UIStrings.EN;
                    break;
            }

            return text;
        };

        return UIStrings;
    })();
    ```

<span data-ttu-id="6c90a-260">UIStrings.js 资源文件创建对象 **UIStrings**，其中包含加载项 UI 的本地化字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-260">The UIStrings.js resource file creates an object, **UIStrings**, which contains the localized strings for your add-in UI.</span></span>

### <a name="localize-the-text-used-for-the-add-in-ui"></a><span data-ttu-id="6c90a-261">本地化加载项 UI 文本</span><span class="sxs-lookup"><span data-stu-id="6c90a-261">Localize the text used for the add-in UI</span></span>

<span data-ttu-id="6c90a-p139">若要在加载项中使用资源文件，需要在 Home.html 中为它添加脚本标记。在 Home.html 加载后，UIStrings.js 便会执行，同时用于获取字符串的 **UIStrings** 对象也可供代码使用。在 Home.html 的头标记中添加以下 HTML，让 **UIStrings** 可供代码使用。</span><span class="sxs-lookup"><span data-stu-id="6c90a-p139">To use the resource file in your add-in, you'll need to add a script tag for it on Home.html. When Home.html is loaded, UIStrings.js executes and the **UIStrings** object that you use to get the strings is available to your code. Add the following HTML in the head tag for Home.html to make **UIStrings** available to your code.</span></span>

```html
<!-- Resource file for localized strings: -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

<span data-ttu-id="6c90a-265">现在，可以使用 **UIStrings** 对象，为加载项 UI 设置字符串了。</span><span class="sxs-lookup"><span data-stu-id="6c90a-265">Now you can use the **UIStrings** object to set the strings for the UI of your add-in.</span></span>

<span data-ttu-id="6c90a-266">如果要根据在 Office 客户端应用程序中的菜单和命令中显示所使用的语言来更改外接程序的本地化，请使用 **Office.context.displayLanguage** 属性获取该语言区域设置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-266">If you want to change the localization for your add-in based on what language is used for display in menus and commands in the Office client application, you use the **Office.context.displayLanguage** property to get the locale for that language.</span></span> <span data-ttu-id="6c90a-267">例如，如果应用程序语言使用西班牙语在菜单和命令中显示 **，Office.context.displayLanguage** 属性将返回语言代码 es-ES。</span><span class="sxs-lookup"><span data-stu-id="6c90a-267">For example, if the application language uses Spanish for display in menus and commands, the **Office.context.displayLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="6c90a-268">如果要根据编辑文档内容所使用的语言更改加载项的本地化，请使用 **Office.context.contentLanguage** 属性获取该语言区域设置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-268">If you want to change the localization for your add-in based on what language is being used for editing document content, you use the **Office.context.contentLanguage** property to get the locale for that language.</span></span> <span data-ttu-id="6c90a-269">例如，如果应用程序语言使用西班牙语编辑文档内容 **，Office.context.contentLanguage** 属性将返回语言代码 es-ES。</span><span class="sxs-lookup"><span data-stu-id="6c90a-269">For example, if the application language uses Spanish for editing document content, the **Office.context.contentLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="6c90a-270">了解应用程序使用的语言后，可以使用 **UIStrings** 获取与应用程序语言匹配的一组本地化字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-270">After you know the language the application is using, you can use **UIStrings** to get the set of localized strings that matches the application language.</span></span>

<span data-ttu-id="6c90a-271">用以下代码替换 Home.js 文件中的代码。</span><span class="sxs-lookup"><span data-stu-id="6c90a-271">Replace the code in the Home.js file with the following code.</span></span> <span data-ttu-id="6c90a-272">该代码演示如何根据应用程序的显示语言或应用程序的编辑语言更改 Home.html 上的 UI 元素中使用的字符串。</span><span class="sxs-lookup"><span data-stu-id="6c90a-272">The code shows how you can change the strings used in the UI elements on Home.html based on either the display language of the application or the editing language of the application.</span></span>

> [!NOTE]
> <span data-ttu-id="6c90a-273">要根据编辑所使用的语言在更改加载项本地化之间进行切换，请取消注释代码行 `var myLanguage = Office.context.contentLanguage;` 并注释掉代码行 `var myLanguage = Office.context.displayLanguage;`</span><span class="sxs-lookup"><span data-stu-id="6c90a-273">To switch between changing the localization of the add-in based on the language used for editing, uncomment the line of code  `var myLanguage = Office.context.contentLanguage;` and comment out the line of code `var myLanguage = Office.context.displayLanguage;`</span></span>

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {

        $(document).ready(function () {
            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // var myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the Office application.
            var myLanguage = Office.context.displayLanguage;
            var UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Introduction);
        });
    };
})();
```

### <a name="test-your-localized-add-in"></a><span data-ttu-id="6c90a-274">测试本地化的加载项</span><span class="sxs-lookup"><span data-stu-id="6c90a-274">Test your localized add-in</span></span>

<span data-ttu-id="6c90a-275">若要测试本地化的外接程序，请更改用于在应用程序Office或编辑的语言，然后运行您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="6c90a-275">To test your localized add-in, change the language used for display or editing in the Office application and then run your add-in.</span></span>

<span data-ttu-id="6c90a-276">若要更改加载项中的显示或编辑语言，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6c90a-276">To change the language used for display or editing in your add-in:</span></span>

1. <span data-ttu-id="6c90a-277">在 Word 中，选择“文件” > “选项” > “语言”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-277">In Word, choose **File** > **Options** > **Language**.</span></span> <span data-ttu-id="6c90a-278">下图显示打开了“语言”选项卡的“Word 选项”对话框。</span><span class="sxs-lookup"><span data-stu-id="6c90a-278">The following figure shows the **Word Options** dialog box opened to the Language tab.</span></span>

    <span data-ttu-id="6c90a-279">*图 2：“Word 选项”对话框中的“语言”选项*</span><span class="sxs-lookup"><span data-stu-id="6c90a-279">*Figure 2. Language options in the Word Options dialog box*</span></span>

    !["Word 选项"对话框。](../images/office15-app-how-to-localize-fig04.png)

2. <span data-ttu-id="6c90a-281">在“**选择显示语言**”下，选择想要显示的语言，例如西班牙语，然后选择向上箭头键将西班牙语移至列表中的第一个位置。</span><span class="sxs-lookup"><span data-stu-id="6c90a-281">Under **Choose Display Language**, select the language that you want for display, for example Spanish, and then choose the up arrow to move the Spanish language to the first position in the list.</span></span> <span data-ttu-id="6c90a-282">或者，若要更改用于编辑的语言，在"选择编辑语言"下，选择要用于编辑的语言，例如西班牙语，然后选择"**设置为默认值"。**</span><span class="sxs-lookup"><span data-stu-id="6c90a-282">Alternatively, to change the language used for editing, under **Choose Editing Languages**, choose the language you want to use for editing, for example, Spanish, and then choose **Set as Default**.</span></span>

3. <span data-ttu-id="6c90a-283">选择“确定”确认选择，然后关闭 Word。</span><span class="sxs-lookup"><span data-stu-id="6c90a-283">Choose **OK** to confirm your selection, and then close Word.</span></span>

4. <span data-ttu-id="6c90a-284">在 Visual Studio 中按 **F5** 以运行示例加载项，或者从菜单栏中选择“调试” > “开始调试”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-284">Press **F5** in Visual Studio to run the sample add-in, or choose **Debug** > **Start Debugging** from the menu bar.</span></span>

5. <span data-ttu-id="6c90a-285">在 Word 中选择“开始” > “显示任务窗格”。</span><span class="sxs-lookup"><span data-stu-id="6c90a-285">In Word, choose **Home** > **Show Taskpane**.</span></span>

<span data-ttu-id="6c90a-286">运行后，外接程序 UI 中的字符串会更改以匹配应用程序使用的语言，如下图所示。</span><span class="sxs-lookup"><span data-stu-id="6c90a-286">Once running, the strings in the add-in UI change to match the language used by the application, as shown in the following figure.</span></span>


<span data-ttu-id="6c90a-287">*图 3. 使用本地化文本的加载项 UI*</span><span class="sxs-lookup"><span data-stu-id="6c90a-287">*Figure 3. Add-in UI with localized text*</span></span>

![具有本地化 UI 文本的应用程序。](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a><span data-ttu-id="6c90a-289">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6c90a-289">See also</span></span>

- [<span data-ttu-id="6c90a-290">Office 加载项的设计准则</span><span class="sxs-lookup"><span data-stu-id="6c90a-290">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- <span data-ttu-id="6c90a-291">[Office 2013 中的语言标识符和 OptionState Id 值](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="6c90a-291">[Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span></span>

[DefaultLocale]:         ../reference/manifest/defaultlocale.md
[说明]:           ../reference/manifest/description.md
[Description]:           ../reference/manifest/description.md
[DisplayName]:           ../reference/manifest/displayname.md
[IconUrl]:               ../reference/manifest/iconurl.md
[HighResolutionIconUrl]: ../reference/manifest/highresolutioniconurl.md
[Resources]:             ../reference/manifest/resources.md
[SourceLocation]:        ../reference/manifest/sourcelocation.md
[Override]:              ../reference/manifest/override.md
[DesktopSettings]:       ../reference/manifest/desktopsettings.md
[TabletSettings]:        ../reference/manifest/tabletsettings.md
[PhoneSettings]:         ../reference/manifest/phonesettings.md
[displayLanguage]:       /javascript/api/office/office.context#displaylanguage
[contentLanguage]:       /javascript/api/office/office.context#contentlanguage
[RFC 3066]:              https://www.rfc-editor.org/info/rfc3066
