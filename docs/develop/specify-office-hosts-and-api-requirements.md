---
title: 指定 Office 主机和 API 要求
description: 了解如何指定 Office 应用程序和 API 要求，使加载项按预期工作。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 948e86e99150ebf2d0bc7deaa5512627679b025f
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237838"
---
# <a name="specify-office-applications-and-api-requirements"></a>指定 Office 应用程序和 API 要求

Office 外接程序可能依赖于特定的 Office 应用程序、要求集、API 成员或 API 版本才能按预期工作。 例如，你的外接程序可能：

- 在单个 Office 应用程序（如，Word 或 Excel）或多个应用程序中运行。

- 使用仅在 Office 的某些版本中可用的 JavaScript API。例如，可能会在运行在 Excel 2016 中的外接程序中使用适用于 Excel 的 JavaScript API。

- 只能在 Office 的某些版本中运行，这些版本支持供你的外接程序使用的 API 成员。

本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。

> [!NOTE]
> 有关 Office 外接程序当前受支持位置的高级别视图，请参阅 Office 外接程序的 Office 客户端应用程序和平台 [可用性](../overview/office-add-in-availability.md) 页面。

下表列出了本文中讨论的核心概念。

|**概念**|**说明**|
|:-----|:-----|
|Office 应用程序、Office 客户端应用程序|用于运行加载项的 Office 应用程序。例如 Word、Excel 等。|
|平台|Office 应用程序的运行位置，例如浏览器或 iPad。|
|要求集|命名的一组相关的 API 成员。 外接程序使用要求集来确定 Office 应用程序是否支持外接程序使用的 API 成员。 测试对要求集的支持比对单个的 API 成员的支持更为容易。 要求集支持因 Office 应用程序和 Office 应用程序的版本而异。 <br >要求集在清单文件中指定。 在清单中指定要求集时，应设置 Office 应用程序为运行外接程序而必须提供的最低级别的 API 支持。 不支持清单中指定的要求集的 Office 应用程序无法运行您的外接程序，并且您的外接程序不会显示在"我的外接程序 <span class="ui">"中</span>。这将限制加载项的可用位置。 在使用运行时检查的代码中。 有关要求集的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。|
|运行时检查|在运行时执行的一个测试，用于确定运行加载项的 Office 应用程序是否支持加载项使用的要求集或方法。 若要执行运行时检查，请使用 **if** 语句以及不是要求集一部分的方法、要求集 `isSetSupported` 或方法名称。 使用运行时检查可确保加载项能够覆盖最大数量的客户。 与要求集不同，运行时检查不会指定 Office 应用程序必须为外接程序运行提供的最低级别的 API 支持。 相反，使用 **if** 语句来确定 API 成员是否受支持。 如果支持，则可以在外接程序中提供其他功能。 使用运行时检查时，你的外接程序将始终在“**我的外接程序**”中显示。|

## <a name="before-you-begin"></a>开始之前

您的外接程序必须使用最新版本的外接程序清单架构。 如果在外接程序中使用运行时检查，请确保使用最新的 Office JavaScript API (office.js) 库。

### <a name="specify-the-latest-add-in-manifest-schema"></a>指定最新的外接程序清单架构

外接程序清单必须使用外接程序清单架构版本 1.1。 按 `OfficeApp` 如下方式设置外接程序清单中的元素。

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>指定最新的 Office JavaScript API 库

如果使用运行时检查，请从 CDN 服务器的内容交付网络引用 Office JavaScript API 库 (最新版本) 。 若要执行此操作，请将以下 `script` 标记添加到 HTML 中。 使用 CDN URL 中的 `/1/` 可以确保引用的是最新版本的 Office.js。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>用于指定 Office 应用程序或 API 要求的选项

指定 Office 应用程序或 API 要求时，有几个因素需要考虑。 下图显示了如何确定要在外接程序中使用的技术。

![指定 Office 应用程序或 API 要求时，为外接程序选择最佳选项](../images/options-for-office-hosts.png)

- 如果加载项在一个 Office 应用程序中运行，请设置 `Hosts` 清单中的元素。 有关详细信息，请参阅 [设置 Hosts 元素](#set-the-hosts-element)。

- 若要设置 Office 应用程序运行加载项时必须支持的最低要求集或 API 成员，请设置 `Requirements` 清单中的元素。 有关详细信息，请参阅[在清单中设置 Requirements 元素](#set-the-requirements-element-in-the-manifest)。

- 如果要在 Office 应用程序中提供特定要求集或 API 成员时提供其他功能，请在加载项的 JavaScript 代码中执行运行时检查。 例如，如果加载项在 Excel 2016 中运行，请使用 Excel JavaScript API 中的 API 成员提供附加功能。 有关详细信息，请参阅[在你的 JavaScript 代码中使用运行时检查](#use-runtime-checks-in-your-javascript-code)。

## <a name="set-the-hosts-element"></a>设置 Hosts 元素

若要使外接程序在一个 Office 客户端应用程序中运行，请使用 `Hosts` 清单中的 and `Host` 元素。 如果不指定元素，加载项将在 Office 加载项支持的所有 Office 应用程序中 `Hosts` 运行。

例如，以下和声明指定外接程序将适用于任何 Excel 版本，其中包括 `Hosts` Excel 网页版、Windows 版和 `Host` iPad 版。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

该 `Hosts` 元素可以包含一个或多个 `Host` 元素。 该 `Host` 元素指定加载项所需的 Office 应用程序。 `Name`该属性是必需的，可以设置为下列值之一。

| 名称          | Office 客户端应用程序                      |
|:--------------|:----------------------------------------------|
| 数据库      | Access Web App                               |
| 文档      | Word 网页、Windows、Mac、iPad           |
| 邮箱       | Outlook 网页版、Windows、Mac、Android、iOS|
| 演示文稿  | PowerPoint 网页、Windows、Mac、iPad     |
| 项目       | Windows 版 Project                            |
| 工作簿      | Excel 网页、Windows、Mac、iPad          |

> [!NOTE]
> 该属性 `Name` 指定可以运行加载项的 Office 客户端应用程序。 Office 应用程序在不同的平台上受支持，并且运行在台式机、Web 浏览器、平板电脑和移动设备上。 不能指定用于运行外接程序的平台。 例如，如果指定，则 Web 上的 Outlook 和 Windows 上的 Outlook 均可用于 `Mailbox` 运行加载项。

> [!IMPORTANT]
> 我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。 作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。

## <a name="set-the-requirements-element-in-the-manifest"></a>在清单中设置 Requirements 元素

该元素指定运行外接程序时 Office 应用程序必须支持的最低要求集 `Requirements` 或 API 成员。 该 `Requirements` 元素可以指定要求集和外接程序中使用的单个方法。 在外接程序清单架构的版本 1.1 中，该元素对于除 Outlook 外接程序之外的所有加载项 `Requirements` 都是可选的。

> [!WARNING]
> 仅使用此元素指定加载项必须使用的关键要求集 `Requirements` 或 API 成员。 如果 Office 应用程序或平台不支持元素中指定的要求集或 API 成员，加载项将不会在应用程序或平台中运行，也不会显示在"我的外接程序 `Requirements` **"中**。相反，我们建议你在 Office 应用程序的所有平台上（如 Web 上的 Excel、Windows 和 iPad）上提供加载项。 若要使加载项在所有  _Office_ 应用程序和平台上可用，请使用运行时检查而不是 `Requirements` 元素。

以下代码示例显示了一个外接程序，该加载项在支持以下功能的所有 Office 客户端应用程序中加载：

-  `TableBindings` 要求集，最低版本为"1.1"。

-  `OOXML` 要求集，最低版本为"1.1"。

-  `Document.getSelectedDataAsync` 方法。

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- 元素 `Requirements` 包含 `Sets` 和 `Methods` 子元素。

- 该 `Sets` 元素可以包含一个或多个 `Set` 元素。 `DefaultMinVersion` 指定所有 `MinVersion` 子元素的 `Set` 默认值。

- 该元素指定 Office 应用程序必须支持的要求集以 `Set` 运行外接程序。 `Name`该属性指定要求集的名称。 `MinVersion`指定要求集的最低版本。 `MinVersion` 替代 API 成员所属的要求集和要求集版本详细信息，请参阅 `DefaultMinVersion` [Office 外接程序要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。

- 该 `Methods` 元素可以包含一个或多个 `Method` 元素。无法将元素 `Methods` 与 Outlook 外接程序一同使用。

- 该元素指定在运行加载项的 Office 应用程序中必须支持的 `Method` 单个方法。该属性是必需的，并指定使用其父 `Name` 对象限定的方法的名称。

## <a name="use-runtime-checks-in-your-javascript-code"></a>在你的 JavaScript 代码中使用运行时检查

如果 Office 应用程序支持某些要求集，您可能需要在外接程序中提供其他功能。 例如，如果你的加载项在 Word 2016 中运行，则你可能想要在现有的加载项中使用 Word JavaScript API。 若要执行此操作，你可以使用含有要求集名称的 [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) 方法。 `isSetSupported` 在运行时确定运行外接程序的 Office 应用程序是否支持要求集。 如果支持要求集，则返回 true 并运行使用该要求集 `isSetSupported` 的 API 成员的其他代码。 如果 Office 应用程序不支持要求集，则返回 `isSetSupported` **false，** 其他代码将不会运行。 以下代码显示要用于的语法 `isSetSupported` 。

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_（必填）是代表该要求集名称的字符串（例如，“**ExcelApi**”、“**Mailbox**”等）。 有关可用要求集的详细信息，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。
- _MinimumVersion_ (可选) 是一个字符串，指定 Office 应用程序必须支持的最低要求集版本，以便语句中的代码运行 (例如 `if` **"1.9**") 。

> [!WARNING]
> 调用该方法 `isSetSupported` 时，参数的值 `MinimumVersion` (指定参数) 字符串。 这是因为 JavaScript 分析器无法区分数值，例如 1.1 和 1.10，因为它可以用于字符串值，例如“1.1”和“1.10”。
> `number` 重载已弃用。

与 `isSetSupported` Office `RequirementSetName` 应用程序关联的使用，如下所示。

|Office 应用程序|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|Word|WordApi|

这些应用程序的方法和要求集在 CDN 上Office.js `isSetSupported` 文件提供。 如果不从 CDN 使用Office.js，加载项可能会生成异常，因为 `isSetSupported` 将未定义。 有关详细信息，请参阅"[指定最新的 Office JavaScript API 库"。](#specify-the-latest-office-javascript-api-library)

以下代码示例演示外接程序如何为可能支持不同要求集或 API 成员的不同 Office 应用程序提供不同的功能。

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
    // Run code that provides additional functionality using the Word JavaScript API when the add-in runs in Word 2016 or later.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run additional code when the Office application is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>使用不属于要求集的方法的运行时检查

部分 API 成员不属于要求集 这仅适用于[属于 Office JavaScript API](../reference/javascript-api-for-office.md)命名空间的 API 成员 (除了 Outlook 邮箱 API) 之外的任何内容，但不包括属于) 中的 Word JavaScript API (任何内容、) 中的 Excel JavaScript API (任何内容或) 命名空间中的 `Office.` [](/javascript/api/outlook)[](../reference/overview/word-add-ins-reference-overview.md) `Word.` [](../reference/overview/excel-add-ins-reference-overview.md) `Excel.` [OneNote JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) `OneNote.` (任何内容。 当加载项依赖于不是要求集一部分的方法时，可以使用运行时检查来确定 Office 应用程序是否支持该方法，如以下代码示例所示。 有关不属于要求集的方法的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)。

> [!NOTE]
> 建议限制在加载项代码中使用此类型运行时检查。

下面的代码示例检查 Office 应用程序是否支持 `document.setSelectedDataAsync` 。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```


## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](add-in-manifests.md)
- [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
