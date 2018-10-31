---
title: Office 外接程序的本地化
description: 可以使用 适用于 Office 的 JavaScript API 来确定区域设置并根据主机应用程序的区域设置显示字符串，或根据数据的区域设置解释或显示数据。
ms.date: 01/23/2018
ms.openlocfilehash: 6271010a08266c71d0f8242acf22cc7b1c730381
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506054"
---
# <a name="localization-for-office-add-ins"></a><span data-ttu-id="d07db-103">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="d07db-103">Localization for Office Add-ins</span></span>

<span data-ttu-id="d07db-p101">可以实施适合Office外接程序的任何本地化方案。Office 外接程序平台的 JavaScript API 和清单架构提供了一些选择。可以使用 Office 的 JavaScript API 确定区域设置并根据主机应用程序的区域设置显示字符串，或根据数据的区域设置解释或显示数据。可以使用清单指定区域设置特定的外接程序文件位置和描述性信息。或者，也可以使用 Microsoft Ajax 脚本来支持全球化和本地化。</span><span class="sxs-lookup"><span data-stu-id="d07db-p101">You can implement any localization scheme that's appropriate for your Office Add-in. The JavaScript API and manifest schema of the Office Add-ins platform provide some choices. You can use the JavaScript API for Office to determine a locale and display strings based on the locale of the host application, or to interpret or display data based on the locale of the data. You can use the manifest to specify locale-specific add-in file location and descriptive information. Alternatively, you can use Microsoft Ajax script to support globalization and localization.</span></span>

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a><span data-ttu-id="d07db-109">使用 JavaScript API 确定区域设置特定的字符串</span><span class="sxs-lookup"><span data-stu-id="d07db-109">Use the JavaScript API to determine locale-specific strings</span></span>

<span data-ttu-id="d07db-110">适用于 Office 的 JavaScript API 提供两个属性，支持显示或解释与主机应用程序和数据的区域设置一致的值：</span><span class="sxs-lookup"><span data-stu-id="d07db-110">The JavaScript API for Office provides two properties that support displaying or interpreting values consistent with the locale of the host application and data:</span></span>

- <span data-ttu-id="d07db-p102">[Context.displayLanguage][displayLanguage] 指定主机应用程序用户界面的区域设置（或语言）。以下示例验证主机应用程序是否使用 en-US 或 fr-Fr 区域设置，并显示特定区域设置的问候语。</span><span class="sxs-lookup"><span data-stu-id="d07db-p102">[Context.displayLanguage][displayLanguage] specifies the locale (or language) of the user interface of the host application. The following example verifies if the host application uses the en-US or fr-FR locale, and displays a locale-specific greeting.</span></span>
    
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

- <span data-ttu-id="d07db-p103">[Context.contentLanguage][contentLanguage] 指定数据的区域设置（或语言）。展开上一个代码示例，不检查 [displayLanguage] 属性，而是将 `myLanguage` 分配给 [contentLanguage] 属性，并使用相同代码的其余部分根据数据的区域设置显示问候语：</span><span class="sxs-lookup"><span data-stu-id="d07db-p103">[Context.contentLanguage][contentLanguage] specifies the locale (or language) of the data. Extending the last code sample, instead of checking the [displayLanguage] property, assign `myLanguage` to the [contentLanguage] property, and use the rest of the same code to display a greeting based on the locale of the data:</span></span>
    
    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a><span data-ttu-id="d07db-115">通过清单控制本地化</span><span class="sxs-lookup"><span data-stu-id="d07db-115">Control localization from the manifest</span></span>


<span data-ttu-id="d07db-p104">每个 Office 外接程序在其清单中指定一个 [DefaultLocale] 元素和区域设置。默认情况下，Office 外接程序平台和 Office 主机应用程序将 [Description]、[DisplayName]、[IconUrl]、[HighResolutionIconUrl] 和 [SourceLocation] 元素的值应用于所有的区域设置。可以通过为每个其他区域设置的上述五个元素中的任意一个指定 [Override] 子元素来选择支持将特定值用于特定的区域设置。[DefaultLocale] 元素和 [Override] 元素的 `Locale` 属性的值根据 [RFC 3066]（“用于语言标识的标记”）指定。表 1 描述了这些元素的本地化支持。</span><span class="sxs-lookup"><span data-stu-id="d07db-p104">Every Office Add-in specifies a [DefaultLocale] element and a locale in its manifest. By default, the Office Add-in platform and Office host applications apply the values of the [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl], and [SourceLocation] elements to all locales. You can optionally support specific values for specific locales, by specifying an [Override] child element for each additional locale, for any of these five elements. The value for the [DefaultLocale] element and for the `Locale` attribute of the [Override] element is specified according to [RFC 3066], "Tags for the Identification of Languages." Table 1 describes the localizing support for these elements.</span></span>

<span data-ttu-id="d07db-121">**表 1.本地化支持**</span><span class="sxs-lookup"><span data-stu-id="d07db-121">**Table 1. Localization support**</span></span>


|<span data-ttu-id="d07db-122">**元素**</span><span class="sxs-lookup"><span data-stu-id="d07db-122">**Element**</span></span>|<span data-ttu-id="d07db-123">**本地化支持**</span><span class="sxs-lookup"><span data-stu-id="d07db-123">**Localization support**</span></span>|
|:-----|:-----|
|<span data-ttu-id="d07db-124">[说明]</span><span class="sxs-lookup"><span data-stu-id="d07db-124">[Description]</span></span>   |<span data-ttu-id="d07db-125">指定的每个区域设置中的用户都可以在 AppSource（或专有目录）中看到本地化的外接程序说明。</span><span class="sxs-lookup"><span data-stu-id="d07db-125">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="d07db-126">对于 Outlook 外接程序，在安装后，用户可以在 Exchange 管理中心 (EAC) 中看到说明。</span><span class="sxs-lookup"><span data-stu-id="d07db-126">For Outlook add-ins, users can see the description in the Exchange Admin Center (EAC) after installation.</span></span>|
|<span data-ttu-id="d07db-127">[DisplayName]</span><span class="sxs-lookup"><span data-stu-id="d07db-127">[DisplayName]</span></span>   |<span data-ttu-id="d07db-128">指定的每个区域设置中的用户都可以在 AppSource（或专有目录）中看到本地化的外接程序说明。</span><span class="sxs-lookup"><span data-stu-id="d07db-128">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="d07db-129">对于 Outlook 外接程序，在安装后，用户可以看到显示名称为 Outlook 外接程序按钮标签，也可以在 EAC 中看到显示名称。</span><span class="sxs-lookup"><span data-stu-id="d07db-129">For Outlook add-ins, users can see the display name as a label for the Outlook add-in button and in the EAC after installation.</span></span><br/><span data-ttu-id="d07db-130">对于内容和任务窗格外接程序，安装外接程序后，用户可以在功能区中看到该显示名称。</span><span class="sxs-lookup"><span data-stu-id="d07db-130">For content and task pane add-ins, users can see the display name in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="d07db-131">[IconUrl]</span><span class="sxs-lookup"><span data-stu-id="d07db-131">[IconUrl]</span></span>        |<span data-ttu-id="d07db-p105">图标图像是可选的。可以使用相同的替代方法为特定区域性指定特定图像。如果使用并本地化图标，则指定的每个区域设置中的用户均可看到该外接程序的本地化图标图像。</span><span class="sxs-lookup"><span data-stu-id="d07db-p105">The icon image is optional. You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="d07db-135">对于 Outlook 外接程序，安装外接程序后，用户可以在 EAC 中看到该图标。</span><span class="sxs-lookup"><span data-stu-id="d07db-135">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="d07db-136">对于内容和任务窗格的外接程序，安装外接程序后，用户可以在功能区中看到此图标。</span><span class="sxs-lookup"><span data-stu-id="d07db-136">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="d07db-137">[HighResolutionIconUrl] **重要说明：** 此元素仅适用于使用版本 1.1 外接程序清单的情况。</span><span class="sxs-lookup"><span data-stu-id="d07db-137">[HighResolutionIconUrl] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>|<span data-ttu-id="d07db-p106">高分辨率图标图像是可选的，但一旦指定，则必须在  [IconUrl] 元素之后出现。指定 [HighResolutionIconUrl] 且在支持高 DPI 分辨率的设备上安装了外接程序后，将使用 [HighResolutionIconUrl] 值而不是 [IconUrl] 值。</span><span class="sxs-lookup"><span data-stu-id="d07db-p106">The high resolution icon image is optional but if it is specified, it must occur after the  [IconUrl] element. When [HighResolutionIconUrl] is specified, and the add-in is installed on a device that supports high dpi resolution, the [HighResolutionIconUrl] value is used instead of the value for [IconUrl].</span></span><br/><span data-ttu-id="d07db-p107">图标图像是可选的。可以使用相同的替代技术为特定区域性指定特定图像。如果使用并本地化图标，则指定的每个区域设置中的用户均可看到该外接程序的本地化图标图像。</span><span class="sxs-lookup"><span data-stu-id="d07db-p107">You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="d07db-142">对于 Outlook 外接程序，安装外接程序后，用户可以在 EAC 中看到该图标。</span><span class="sxs-lookup"><span data-stu-id="d07db-142">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="d07db-143">对于内容和任务窗格的外接程序，安装外接程序后，用户可以在功能区中看到此图标。</span><span class="sxs-lookup"><span data-stu-id="d07db-143">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="d07db-144">[Resources] **重要说明：** 此元素仅适用于使用版本 1.1 外接程序清单的情况。</span><span class="sxs-lookup"><span data-stu-id="d07db-144">[Resources] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>   |<span data-ttu-id="d07db-145">指定的每个区域设置中的用户都可以看到专门为该区域设置的外接程序创建的字符串和图标资源 。</span><span class="sxs-lookup"><span data-stu-id="d07db-145">Users in each locale you specify can see string and icon resources that you specifically create for the add-in for that locale.</span></span> |
|<span data-ttu-id="d07db-146">[SourceLocation]</span><span class="sxs-lookup"><span data-stu-id="d07db-146">[SourceLocation]</span></span>   |<span data-ttu-id="d07db-147">指定的每个区域设置中的用户都可以看到回复专门为该区域设置的外接程序创建的网页。</span><span class="sxs-lookup"><span data-stu-id="d07db-147">Users in each locale you specify can see a webpage that you specifically design for the add-in for that locale.</span></span> |


> [!NOTE] 
> <span data-ttu-id="d07db-p108">只能本地化 Office 支持的区域的说明和显示名称。欲知最新版 Office 支持的语言和区域设置列表，请参阅 [Office 2013 中的语言标识符和 OptionState Id 值](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))。</span><span class="sxs-lookup"><span data-stu-id="d07db-p108">You can localize the description and display name for only the locales that Office supports. See [Language identifiers and OptionState Id values in Office 2013](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) for a list of languages and locales for the current release of Office.</span></span>


### <a name="examples"></a><span data-ttu-id="d07db-150">示例</span><span class="sxs-lookup"><span data-stu-id="d07db-150">Examples</span></span>

<span data-ttu-id="d07db-p109">例如，Office 外接程序可以将 [DefaultLocale] 指定为 `en-us`。对于 [DisplayName] 元素，外接程序可以为区域设置 `fr-fr` 指定 [Override] 子元素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="d07db-p109">For example, an Office Add-in can specify the [DefaultLocale] as `en-us`. For the [DisplayName] element, the add-in can specify an [Override] child element for the locale `fr-fr`, as shown below.</span></span> 


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE] 
> <span data-ttu-id="d07db-p110">注意：如需针对一个语系内的多个区域（如 `de-de` 和 `de-at`）进行本地化，建议对各个区域使用独立的 `Override` 元素。在这种情况下，仅使用语言名称，`de` ，在所有 Office主机应用程序和平台组合中都不受支持。</span><span class="sxs-lookup"><span data-stu-id="d07db-p110">If you need to localize for more than one area within a language family, such as `de-de` and `de-at`, we recommend that you use separate `Override` elements for each area. Using just the language name alone, in this case, `de`, is not supported across all combinations of Office host applications and platforms.</span></span>

<span data-ttu-id="d07db-p111">这意味着，外接程序默认情况下采用 `en-us` 区域设置。除非客户端计算机的区域设置为 `fr-fr`（此时用户将看到法语的显示名称 “Lecteur vidéo” ），否则对于所有区域设置，用户都将看到英文显示名称 “Video player” 。</span><span class="sxs-lookup"><span data-stu-id="d07db-p111">This means that the add-in assumes the  `en-us` locale by default. Users see the English display name of "Video player" for all locales unless the client computer's locale is `fr-fr`, in which case users would see the French display name "Lecteur vidéo".</span></span>

> [!NOTE] 
> <span data-ttu-id="d07db-p112">只能为每种语言指定一个覆盖，包括于默认区域设置。例如，如果您的默认区域设置是`en-us` ，则无法为`en-us` 指定覆盖。</span><span class="sxs-lookup"><span data-stu-id="d07db-p112">You may only specify a single override per language, including for the default locale. For example, if your default locale is `en-us` you cannot not specify an  override for `en-us` as well.</span></span> 

<span data-ttu-id="d07db-p113">下面的示例对 [Description] 元素应用区域设置覆盖。它先指定默认区域设置 `en-us` 和英文说明，再指定 [Override] 语句，其中包含 `fr-fr` 区域设置的法语说明：</span><span class="sxs-lookup"><span data-stu-id="d07db-p113">The following example applies a locale override for the [Description] element. It first specifies a default locale of `en-us` and an English description, and then specifies an [Override] statement with a French description for the `fr-fr` locale:</span></span>

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive 
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook et Outlook Web App."/>
</Description>
```

<span data-ttu-id="d07db-p114">也就是说，外接程序默认采用 `en-us` 区域设置。除非客户端计算机的区域设置为 `fr-fr`（此时用户将看到法语说明），否则对于所有区域设置，用户都将看到 `DefaultValue` 属性中的英文说明。</span><span class="sxs-lookup"><span data-stu-id="d07db-p114">This means that the add-in assumes the `en-us` locale by default. Users would see the English description in the `DefaultValue` attribute for all locales unless the client computer's locale is `fr-fr`, in which case they would see the French description.</span></span>

<span data-ttu-id="d07db-p115">在以下示例中，外接程序指定更适合 `fr-fr` 区域设置和区域性的不同图像。默认情况下，用户看到的是图像 DefaultLogo.png，客户端计算机的区域设置为 `fr-fr` 时除外。这种情况下，用户将看到图像 FrenchLogo.png。</span><span class="sxs-lookup"><span data-stu-id="d07db-p115">In the following example, the add-in specifies a separate image that's more appropriate for the `fr-fr` locale and culture. Users see the image DefaultLogo.png by default, except when the locale of the client computer is `fr-fr`. In this case, users would see the image FrenchLogo.png.</span></span> 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

<span data-ttu-id="d07db-p116">以下示例显示了如何本地化 `Resources` 部分中的资源。它对一个更适用于 `ja-jp` 区域性的图像应用了区域设置覆盖。</span><span class="sxs-lookup"><span data-stu-id="d07db-p116">The following example shows how to localize a resource in the `Resources` section. It applies a locale override for an image that is more appropriate for the `ja-jp` culture.</span></span>

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


<span data-ttu-id="d07db-p117">关于 [SourceLocation] 元素，支持其他区域设置意味着为每个指定的区域设置提供单独的源 HTML 文件。指定的每个区域设置中的用户都可以看到为相应区域设置设计的自定义网页。</span><span class="sxs-lookup"><span data-stu-id="d07db-p117">For the [SourceLocation] element, supporting additional locales means providing a separate source HTML file for each of the specified locales. Users in each locale you specify can see a customized webpage that you design for that them.</span></span>

<span data-ttu-id="d07db-p118">对于 Outlook 外接程序，[SourceLocation] 元素还与外形规格保持一致。这样一来，就可以为每个相应外形规格提供不同的本地化源 HTML 文件。可以在每个适用的 settings 元素（[DesktopSettings]、[TabletSettings] 或 [PhoneSettings]）中指定一个或多个 [Override] 子元素。下面的示例展示了用于台式机、平板电脑和智能手机外形规格的 settings 元素，每个都有一个用于默认区域设置的 HTML 文件，以及另一个用于法语区域设置的 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="d07db-p118">For Outlook add-ins, the [SourceLocation] element also aligns to the form factor. This allows you to provide a separate, localized source HTML file for each corresponding form factor. You can specify one or more [Override] child elements in each applicable settings element ([DesktopSettings], [TabletSettings], or [PhoneSettings]). The following example shows settings elements for the desktop, tablet, and smartphone form factors, each with one HTML file for the default locale and another for the French locale.</span></span>


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

## <a name="match-datetime-format-with-client-locale"></a><span data-ttu-id="d07db-174">将日期/时间格式与客户端区域设置匹配</span><span class="sxs-lookup"><span data-stu-id="d07db-174">Match date/time format with client locale</span></span>

<span data-ttu-id="d07db-p119">可以通过 [displayLanguage] 属性获取主机应用程序用户界面的区域位置。然后可以显示格式与主机应用程序中的当前区域位置一致的日期和时间值。执行上述操作的一种方法是准备一个指定日期/时间显示格式的资源文件以用于 Office 外界程序支持的各个区域设置。在运行时，外接程序可以使用该资源文件并匹配通过 [displayLanguage] 获得的区域设置正确的日期/时间格式。</span><span class="sxs-lookup"><span data-stu-id="d07db-p119">You can get the locale of the user interface of the hosting application by using the [displayLanguage] property. You can then display date and time values in a format consistent with the current locale of the host application. One way to do that is to prepare a resource file that specifies the date/time display format to use for each locale that your Office Add-in supports. At run time, your add-in can use the resource file and match the appropriate date/time format with the locale obtained from the [displayLanguage] property.</span></span>

<span data-ttu-id="d07db-p120">可以使用 [contentLanguage] 属性，获取主机应用数据的区域设置。根据此值，可以正确地解释或显示日期/时间字符串。例如，`jp-JP` 区域设置将数据/时间值表示为 `yyyy/MM/dd`，而 `fr-FR` 区域设置则表示为 `dd/MM/yyyy`。</span><span class="sxs-lookup"><span data-stu-id="d07db-p120">You can get the locale of the data of the hosting application by using the [contentLanguage] property. Based on this value, you can then appropriately interpret or display date/time strings. For example, the `jp-JP` locale expresses data/time values as `yyyy/MM/dd`, and the `fr-FR` locale, `dd/MM/yyyy`.</span></span>


## <a name="use-ajax-for-globalization-and-localization"></a><span data-ttu-id="d07db-182">将 Ajax 用于全球化和本地化</span><span class="sxs-lookup"><span data-stu-id="d07db-182">Use Ajax for globalization and localization</span></span>


<span data-ttu-id="d07db-183">如果使用 Visual Studio 创建 Office 外接程序，.NET Framework 和 Ajax 会提供用于全球化和本地化客户端脚本文件的方法。</span><span class="sxs-lookup"><span data-stu-id="d07db-183">If you use Visual Studio to create Office Add-ins, the .NET Framework and Ajax provide ways to globalize and localize client script files.</span></span>

<span data-ttu-id="d07db-p121">您可以全球化 Office 外接程序并在其 JavaScript 代码中使用 [Date](http://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) 和 [Number](http://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) JavaScript 类型扩展和 JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) 对象，以根据当前浏览器的区域设置显示值。欲知详情，请参阅 [Walkthrough: Globalizing a Date by Using Client Script](http://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx)。</span><span class="sxs-lookup"><span data-stu-id="d07db-p121">You can globalize and use the [Date](http://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) and [Number](http://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) JavaScript type extensions and the JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object in the JavaScript code for an Office Add-in to display values based on the locale settings on the current browser. For more information, see [Walkthrough: Globalizing a Date by Using Client Script](http://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).</span></span>

<span data-ttu-id="d07db-p122">可将本地化的资源字符串直接包含在独立的 JavaScript 文件中，以便为不同区域设置提供客户端脚本文件，这些文件在浏览器中设置或由用户提供。为每个受支持的区域设置创建单独的脚本文件。在每个脚本文件中，包括一个 JSON 格式的对象，其中包含用于该区域设置的资源字符串。在浏览器中运行脚本时，会应用本地化的值。</span><span class="sxs-lookup"><span data-stu-id="d07db-p122">You can include localized resource strings directly in standalone JavaScript files to provide client script files for different locales, which are set on the browser or provided by the user. Create a separate script file for each supported locale. In each script file, include an object in JSON format that contains the resource strings for that locale. The localized values are applied when the script runs in the browser.</span></span> 


## <a name="example-build-a-localized-office-add-in"></a><span data-ttu-id="d07db-190">示例：生成本地化 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="d07db-190">Example: Build a localized Office Add-in</span></span>

<span data-ttu-id="d07db-191">本节提供示例，演示如何本地化 Office 外接程序说明、显示名称和 UI。</span><span class="sxs-lookup"><span data-stu-id="d07db-191">This section provides examples that show you how to localize an Office Add-in description, display name, and UI.</span></span>

<span data-ttu-id="d07db-192">若要运行所提供的示例代码，请在计算机上配置 Microsoft Office 2013 以使用其他语言，这样就可以通过切换用于在菜单和命令中显示的语言，编辑和校对，或两者同时来测试外接程序。</span><span class="sxs-lookup"><span data-stu-id="d07db-192">To run the sample code provided, configure Microsoft Office 2013 on your computer to use additional languages so that you can test your add-in by switching the language used for display in menus and commands, for editing and proofing, or both.</span></span>

<span data-ttu-id="d07db-193">此外，还需要创建 Visual Studio 2015 Office 外接程序项目。</span><span class="sxs-lookup"><span data-stu-id="d07db-193">Also, you'll need to create a Visual Studio 2015 Office Add-in project.</span></span>

> [!NOTE] 
> <span data-ttu-id="d07db-p123">若要下载 Visual Studio 2015，请参阅  [Office Developer Tools page](https://www.visualstudio.com/features/office-tools-vs) 。此页面还包含 Office 开发人员工具的链接。</span><span class="sxs-lookup"><span data-stu-id="d07db-p123">To download Visual Studio 2015, see the [Office Developer Tools page](https://www.visualstudio.com/features/office-tools-vs). This page also has a link for the Office Developer Tools.</span></span>

### <a name="configure-office-2013-to-use-additional-languages-for-display-or-editing"></a><span data-ttu-id="d07db-196">配置 Office 2013 使用额外的显示或编辑语言</span><span class="sxs-lookup"><span data-stu-id="d07db-196">Configure Office 2013 to use additional languages for display or editing</span></span>

<span data-ttu-id="d07db-p124">您可以使用 Office 2013 语言包安装其他语言。欲知语言包及其获取位置的详细信息，请参阅 [Office 2013 语言选项](http://office.microsoft.com/language-packs/)。</span><span class="sxs-lookup"><span data-stu-id="d07db-p124">You can use an Office 2013 Language pack to install an additional language. For more information about Language Packs and where to get them, see [Office 2013 Language Options](http://office.microsoft.com/language-packs/).</span></span>

> [!NOTE] 
> <span data-ttu-id="d07db-p125">如果您是一位 MSDN 订户，则可能已拥有适用的 Office 2013 语言包。 若要确定您的订阅是否有可供下载的 Office 2013 语言包，请转到 [  MSDN Subscriptions Home](https://msdn.microsoft.com/subscriptions/manage/)  ，在 \*\* Software downloads\*\* 中输入 Office 2013 语言包 ，选择 \*\* Search\*\* ，然后选择 \*\* Products available with my subscription\*\*  。 在 **Language** 下，选中想要下载的语言包的复选框，然后选择 \*\* Go\*\* 。</span><span class="sxs-lookup"><span data-stu-id="d07db-p125">If you are an MSDN Subscriber, you might already have the Office 2013 Language Packs available to you. To determine whether your subscription offers Office 2013 Language Packs for download, go to [MSDN Subscriptions Home](https://msdn.microsoft.com/subscriptions/manage/), enter Office 2013 Language Pack in **Software downloads**, choose **Search**, and then select **Products available with my subscription**. Under **Language**, select the check box for the Language Pack you want to download, and then choose  **Go**.</span></span> 

<span data-ttu-id="d07db-p126">安装语言包后，可以配置 Office 2013 以使用安装的语言在 UI 中显示或编辑文档内容，或同时用于两者。本文中的示例使用已应用西班牙语语言包的 Office 2013安装。</span><span class="sxs-lookup"><span data-stu-id="d07db-p126">After you install the Language Pack, you can configure Office 2013 to use the installed language for display in the UI, for editing document content, or both. The example in this article uses an installation of Office 2013 that has the Spanish Language Pack applied.</span></span>

### <a name="create-an-office-add-in-project"></a><span data-ttu-id="d07db-204">创建 Office 外接程序项目</span><span class="sxs-lookup"><span data-stu-id="d07db-204">Create an Office Add-in project</span></span>

1. <span data-ttu-id="d07db-205">在 Visual Studio 中，选择\*\*\*\* File > New Project\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="d07db-205">In Visual Studio, choose **File** > **New Project**.</span></span>
    
2. <span data-ttu-id="d07db-206">在 \*\* New Project\*\*  对话框的 \*\* Templates\*\* 下，展开 \*\* Visual Basic\*\*  或 **Visual C＃**  ，展开 **Office / SharePoint** ，然后选择 \*\* Office Add-ins\*\* 。</span><span class="sxs-lookup"><span data-stu-id="d07db-206">In the **New Project** dialog box, under **Templates**, expand **Visual Basic** or **Visual C#**, expand **Office/SharePoint**, and then choose  **Office Add-ins**.</span></span>
    
3. <span data-ttu-id="d07db-p127">选择 **Office Add-ins**，再为外接程序命名，例如，WorldReadyAddIn 。选择 **OK**。</span><span class="sxs-lookup"><span data-stu-id="d07db-p127">Choose **Office Add-in**, and then name your add-in, for example WorldReadyAddIn. Choose  **OK**.</span></span>
    
4. <span data-ttu-id="d07db-p128">在  **Create Office Add-in**  对话框中，选择 **Task pane**  和  **Next** 。在下一个页面上，清除所有主机应用的复选框，Word 除外。选择 **Finish** 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="d07db-p128">In the **Create Office Add-in** dialog box, select **Task pane** and choose **Next**. On the next page, clear the check boxes for all host applications except Word. Choose **Finish** to create the project.</span></span>
    

### <a name="localize-the-text-used-in-your-add-in"></a><span data-ttu-id="d07db-212">本地化外接程序中使用的文本</span><span class="sxs-lookup"><span data-stu-id="d07db-212">Localize the text used in your add-in</span></span>

<span data-ttu-id="d07db-213">想要本地化成另一种语言的文本出现在两个区域中：</span><span class="sxs-lookup"><span data-stu-id="d07db-213">The text that you want to localize for another language appears in two areas:</span></span>

-  <span data-ttu-id="d07db-p129">**加载项显示名称和说明** 。这是受应用程序清单文件中的条目控制的。</span><span class="sxs-lookup"><span data-stu-id="d07db-p129">**Add-in display name and description**. This is controlled by entries in the add-in manifest file.</span></span>
    
-  <span data-ttu-id="d07db-p130">**外接程序 UI** 。可以通过使用 JavaScript 代码本地化在外接程序 UI 中出现的字符串（例如，通过使用包含本地化后字符串的单独资源文件）。</span><span class="sxs-lookup"><span data-stu-id="d07db-p130">**Add-in UI**. You can localize the strings that appear in your add-in UI by using JavaScript codeâ€”for example, by using a separate resource file that contains the localized strings.</span></span>
    
<span data-ttu-id="d07db-218">若要本地化外接程序显示名称和说明，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="d07db-218">To localize the add-in display name and description:</span></span>

1. <span data-ttu-id="d07db-219">在 \*\* Solution Explorer\*\* 中，展开 **WorldReadyAddIn**， **WorldReadyAddInManifest** ，再选择  **WorldReadyAddIn.xml**。</span><span class="sxs-lookup"><span data-stu-id="d07db-219">In **Solution Explorer**, expand **WorldReadyAddIn**, **WorldReadyAddInManifest**, and then choose  **WorldReadyAddIn.xml**.</span></span>
    
2. <span data-ttu-id="d07db-220">在 WorldReadyAddInManifest.xml 中，将“DisplayName”[]和“Description”[]元素替换为以下代码块：</span><span class="sxs-lookup"><span data-stu-id="d07db-220">In WorldReadyAddInManifest.xml, replace the [DisplayName] and [Description] elements with the following block of code:</span></span>
    
    > [!NOTE] 
    > <span data-ttu-id="d07db-221">对于本示例中使用的西班牙语本地化字符串的[DisplayName] 和 [Description] 元素，可以替换为任何其他语言的本地化字符串。</span><span class="sxs-lookup"><span data-stu-id="d07db-221">You can replace the Spanish language localized strings used in this example for the [DisplayName] and [Description] elements with the localized strings for any other language.</span></span>

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. <span data-ttu-id="d07db-222">例如，如果将 Office 2013 的显示语言从英语切换到西班牙语，然后运行外接程序，外接程序的显示名称和说明将用本地化文本显示。</span><span class="sxs-lookup"><span data-stu-id="d07db-222">When you change the display language for Office 2013 from English to Spanish, for example, and then run the add-in, the add-in display name and description are shown with localized text.</span></span> 
    
<span data-ttu-id="d07db-223">若要设计外接程序 UI 的布局，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="d07db-223">To lay out the add-in UI:</span></span>

1. <span data-ttu-id="d07db-224">在 Visual Studio 的 **Solution Explorer**中，选择 **Home.html**。</span><span class="sxs-lookup"><span data-stu-id="d07db-224">In Visual Studio, in **Solution Explorer**, choose **Home.html**.</span></span>
    
2. <span data-ttu-id="d07db-225">将 Home.html 中的 HTML 替换为以下 HTML 。</span><span class="sxs-lookup"><span data-stu-id="d07db-225">Replace the HTML in Home.html with the following HTML.</span></span>
    
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.8.2.js" type="text/javascript"></script>
    
        <link href="../../Content/Office.css" rel="stylesheet" type="text/css" />
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    
        <!-- To enable offline debugging using a local reference to Office.js, use:                        -->
        <!-- <script src="../../Scripts/Office/MicrosoftAjax.js" type="text/javascript"></script>          -->
        <!--    <script src="../../Scripts/Office/1.0/office.js" type="text/javascript"></script>          -->
    
        <link href="../App.css" rel="stylesheet" type="text/css" />
        <script src="../App.js" type="text/javascript"></script>
    
        <link href="Home.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script> <body>
        <!-- Page content -->
        <div id="content-header">
            <div class="padding">
                <h1 id="greeting"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div>
                    <p id="about"></p>
                </div>            
            </div>
        </div>
    </head>
    </html>
    ```

3. <span data-ttu-id="d07db-226">在 Visual Studio 中，依次**File** 和 **Save AddIn\Home\Home.html** 。</span><span class="sxs-lookup"><span data-stu-id="d07db-226">In Visual Studio, choose  **File**,  **Save AddIn\Home\Home.html**.</span></span>
    
<span data-ttu-id="d07db-227">下图展示了将在示例外接程序运行时显示本地化文本的 heading (h1) 元素和 paragraph (p) 元素。</span><span class="sxs-lookup"><span data-stu-id="d07db-227">The following figure shows the heading (h1) element and the paragraph (p) element that will display localized text when your sample add-in runs.</span></span>

<span data-ttu-id="d07db-228">*图 1：外接程序 UI*</span><span class="sxs-lookup"><span data-stu-id="d07db-228">*Figure 1. The add-in UI*</span></span>

![突出显示了各部分的应用用户界面](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a><span data-ttu-id="d07db-230">添加包含本地化字符串的资源文件</span><span class="sxs-lookup"><span data-stu-id="d07db-230">Add the resource file that contains the localized strings</span></span>

<span data-ttu-id="d07db-p131">JavaScript 资源文件包含用于外接程序 UI 的字符串。示例外接程序 UI 有一个用于显示问候语的 h1 元素，和一个用于向用户介绍该外接程序的 p 元素。</span><span class="sxs-lookup"><span data-stu-id="d07db-p131">The JavaScript resource file contains the strings used for the add-in UI. The sample add-in UI has an h1 element that displays a greeting, and a p element that introduces the add-in to the user.</span></span> 

<span data-ttu-id="d07db-p132">要为标题和段落启用本地化字符串，需要将字符串放在一个单独的资源文件中。资源文件会创建一个 JavaScript 对象，对每组本地化字符串来说，它都包含一个单独的 JavaScript 对象表示法 (JSON) 对象。资源文件也提供为给定区域设置找回适当 JSON 对象的方法。</span><span class="sxs-lookup"><span data-stu-id="d07db-p132">To enable localized strings for the heading and paragraph, you place the strings in a separate resource file. The resource file creates a JavaScript object that contains a separate JavaScript Object Notation (JSON) object for each set of localized strings. The resource file also provides a method for getting back the appropriate JSON object for a given locale.</span></span> 

<span data-ttu-id="d07db-236">若要将资源文件添加到外接程序项目，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="d07db-236">To add the resource file to the add-in project:</span></span>

1. <span data-ttu-id="d07db-237">在 Visual Studio 的 **Solution Explorer** 中，选择 Web 项目中示例外接程序的 **Add-in** 文件夹，再依次选择 **Add** > **JavaScript file** 。</span><span class="sxs-lookup"><span data-stu-id="d07db-237">In **Solution Explorer** in Visual Studio, choose the **Add-in** folder in the web project for the sample add-in, and choose **Add** > **JavaScript file**.</span></span>
    
2. <span data-ttu-id="d07db-238">在**Specify Name for Item** 对话框中，输入 UIStrings.js 。</span><span class="sxs-lookup"><span data-stu-id="d07db-238">In the **Specify Name for Item** dialog box, enterUIStrings.js.</span></span>
    
3. <span data-ttu-id="d07db-239">将以下代码添加到 UIStrings.js 文件。</span><span class="sxs-lookup"><span data-stu-id="d07db-239">Add the following code to the UIStrings.js file.</span></span>

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

<span data-ttu-id="d07db-240">UIStrings.js 资源文件创建对象 **UIStrings**，其中包含外接程序 UI 的本地化字符串。</span><span class="sxs-lookup"><span data-stu-id="d07db-240">The UIStrings.js resource file creates an object, **UIStrings**, which contains the localized strings for your add-in UI.</span></span> 

### <a name="localize-the-text-used-for-the-add-in-ui"></a><span data-ttu-id="d07db-241">本地化外接程序 UI 文本</span><span class="sxs-lookup"><span data-stu-id="d07db-241">Localize the text used for the add-in UI</span></span>

<span data-ttu-id="d07db-p133">若要在外接程序中使用资源文件，需要在 Home.html 中为它添加脚本标记。在 Home.html 加载后，UIStrings.js 便会执行，同时用于获取字符串的 **UIStrings** 对象也可供代码使用。在 Home.html 的头标记中添加以下 HTML，让 **UIStrings** 可供代码使用。</span><span class="sxs-lookup"><span data-stu-id="d07db-p133">To use the resource file in your add-in, you'll need to add a script tag for it on Home.html. When Home.html is loaded, UIStrings.js executes and the **UIStrings** object that you use to get the strings is available to your code. Add the following HTML in the head tag for Home.html to make **UIStrings** available to your code.</span></span>

```html
<!-- Resource file for localized strings:                                                          -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

<span data-ttu-id="d07db-245">现在，可以使用 **UIStrings** 对象，为外接程序 UI 设置字符串了。</span><span class="sxs-lookup"><span data-stu-id="d07db-245">Now you can use the **UIStrings** object to set the strings for the UI of your add-in.</span></span>

<span data-ttu-id="d07db-p134">若要根据主机应用中的菜单和命令的显示语言来更改外接程序的本地化，可以使用 **Office.context.displayLanguage** 属性，获取相应语言的区域设置。例如，如果主机应用语言使用西班牙语显示菜单和命令，那么 **Office.context.displayLanguage** 属性返回语言代码 es-ES。</span><span class="sxs-lookup"><span data-stu-id="d07db-p134">If you want to change the localization for your add-in based on what language is used for display in menus and commands in the host application, you use the **Office.context.displayLanguage** property to get the locale for that language. For example, if the host application language uses Spanish for display in menus and commands, the **Office.context.displayLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="d07db-p135">如果要根据编辑文档内容所用的语言更改的外接程序的本地化，可以使用  **Office.context.contentLanguage**  属性获取该语言的区域设置。例如，如果主机应用程序语言使用西班牙语编辑文档内容， **Office.context.contentLanguage**  属性将返回语言代码 es-ES。</span><span class="sxs-lookup"><span data-stu-id="d07db-p135">If you want to change the localization for your add-in based on what language is being used for editing document content, you use the  **Office.context.contentLanguage** property to get the locale for that language. For example, if the host application language uses Spanish for editing document content, the **Office.context.contentLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="d07db-250">确定主机应用使用的语言后，可以使用 **UIStrings**，获取与主机应用语言一致的本地化字符串组。</span><span class="sxs-lookup"><span data-stu-id="d07db-250">After you know the language the host application is using, you can use **UIStrings** to get the set of localized strings that matches the host application language.</span></span>

<span data-ttu-id="d07db-p136">用以下代码替换 Home.js 文件中的代码。该代码显示了如何根据主机应用程序的显示语言或主机应用程序的编辑语言更改 Home.html 上 UI 元素中使用的字符串。</span><span class="sxs-lookup"><span data-stu-id="d07db-p136">Replace the code in the Home.js file with the following code. The code shows how you can change the strings used in the UI elements on Home.html based on either the display language of the host application or the editing language of the host application.</span></span>

> [!NOTE] 
> <span data-ttu-id="d07db-253">若要根据编辑语言切换更改外接程序的本地化，请取消注释代码行 `var myLanguage = Office.context.contentLanguage;`，并注释掉代码行 `var myLanguage = Office.context.displayLanguage;`</span><span class="sxs-lookup"><span data-stu-id="d07db-253">NOTE To switch between changing the localization of the add-in based on the language used for editing, uncomment the line of code  `var myLanguage = Office.context.contentLanguage;` and comment out the line of code `var myLanguage = Office.context.displayLanguage;`</span></span>

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {
       
        $(document).ready(function () {
            app.initialize();

            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // var myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the host application.
            var myLanguage = Office.context.displayLanguage;            
            var UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);            

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Instruction);
        });
    };    
})();
```

### <a name="test-your-localized-add-in"></a><span data-ttu-id="d07db-254">测试本地化的外接程序</span><span class="sxs-lookup"><span data-stu-id="d07db-254">Test your localized add-in</span></span>

<span data-ttu-id="d07db-255">若要测试本地化外接程序，请更改在主机应用程序中用于显示或编辑的语言，然后运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="d07db-255">To test your localized add-in, change the language used for display or editing in the host application and then run your add-in.</span></span> 

<span data-ttu-id="d07db-256">要更改外接程序中的显示或编辑语言，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="d07db-256">To change the language used for display or editing in your add-in:</span></span>

1. <span data-ttu-id="d07db-p137">在 Word 2013 中，依次选择 **File** > **Options** > **Language** 。下图展示了用**Word Options** 对话框打开语言选项卡。</span><span class="sxs-lookup"><span data-stu-id="d07db-p137">In Word 2013, choose **File** > **Options** > **Language**. The following figure shows the **Word Options** dialog box opened to the Language tab.</span></span>
    
    <span data-ttu-id="d07db-259">*图 2 ：Word 2013 选项对话框中的语言选项*</span><span class="sxs-lookup"><span data-stu-id="d07db-259">*Figure 2. Language options in the Word 2013 Options dialog box*</span></span>

    ![Word 2013 选项对话框](../images/office15-app-how-to-localize-fig04.png)

2. <span data-ttu-id="d07db-p138">在 **Choose Display and Help Languages** ，选择要使用的显示语言，例如，西班牙语，再选择向上箭头键将西班牙语移到列表中的首位。或者，要更改编辑语言，在 **Choose editing languages** 下，选择要使用的编辑语言，例如，西班牙语，再选择  **Set as Default** 。</span><span class="sxs-lookup"><span data-stu-id="d07db-p138">Under **Choose Display and Help Languages**, select the language that you want for display, for example Spanish, and then choose the up arrow to move the Spanish language to the first position in the list. Alternatively, to change the language used for editing, under  **Choose editing languages**, choose the language you want to use for editing, for example, Spanish, and then choose **Set as Default**.</span></span>
    
3. <span data-ttu-id="d07db-263">选择**OK** 确认选择，再关闭 Word 。</span><span class="sxs-lookup"><span data-stu-id="d07db-263">Choose **OK** to confirm your selection, and then close Word.</span></span>
    
<span data-ttu-id="d07db-p139">运行示例外接程序。此时，任务窗格外接程序在 Word 2013 中加载，同时外接程序 UI 字符串更改为与主机应用使用的语言一致，如下图所示。</span><span class="sxs-lookup"><span data-stu-id="d07db-p139">Run the sample add-in. The taskpane add-in loads in Word 2013, and the strings in the add-in UI change to match the language used by the host application, as shown in the following figure.</span></span>


<span data-ttu-id="d07db-266">*图 3：包含本地化文本的外接程序 UI*</span><span class="sxs-lookup"><span data-stu-id="d07db-266">*Figure 3. Add-in UI with localized text*</span></span>

![包含本地化 UI 文本的应用](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a><span data-ttu-id="d07db-268">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d07db-268">See also</span></span>

- [<span data-ttu-id="d07db-269">Office 外接程序的设计准则</span><span class="sxs-lookup"><span data-stu-id="d07db-269">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)    
- <span data-ttu-id="d07db-270">[Office 2013 中的语言标识符和 OptionState Id 值](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="d07db-270">[Language identifiers and OptionState Id values in Office 2013](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span></span>

[DefaultLocale]:        https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultlocale?view=office-js
[说明]:          https://docs.microsoft.com/office/dev/add-ins/reference/manifest/description?view=office-js
[Description]:          https://docs.microsoft.com/office/dev/add-ins/reference/manifest/description?view=office-js
[DisplayName]:          https://docs.microsoft.com/office/dev/add-ins/reference/manifest/displayname?view=office-js
[IconUrl]:              https://docs.microsoft.com/office/dev/add-ins/reference/manifest/iconurl?view=office-js
[HighResolutionIconUrl]:https://docs.microsoft.com/office/dev/add-ins/reference/manifest/highresolutioniconurl?view=office-js
[资源]:            https://docs.microsoft.com/office/dev/add-ins/reference/manifest/resources?view=office-js
[Resources]:            https://docs.microsoft.com/office/dev/add-ins/reference/manifest/resources?view=office-js
[SourceLocation]:       https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js
[Override]:             https://docs.microsoft.com/office/dev/add-ins/reference/manifest/override?view=office-js
[DesktopSettings]:      https://docs.microsoft.com/office/dev/add-ins/reference/manifest/desktopsettings?view=office-js
[TabletSettings]:       https://docs.microsoft.com/office/dev/add-ins/reference/manifest/tabletsettings?view=office-js
[PhoneSettings]:        https://docs.microsoft.com/office/dev/add-ins/reference/manifest/phonesettings?view=office-js
[displayLanguage]:  https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#displaylanguage 
[contentLanguage]:  https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#contentlanguage 
[RFC 3066]: https://www.rfc-editor.org/info/rfc3066
