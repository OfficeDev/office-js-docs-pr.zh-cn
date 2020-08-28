---
title: 指定 Office 主机和 API 要求
description: 了解如何指定加载项的 Office 应用程序和 API 要求以按预期方式工作。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 90ee7c3a5ad01252336608c02f995bbcbbe94212
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292627"
---
# <a name="specify-office-applications-and-api-requirements"></a>指定 Office 应用程序和 API 要求

您的 Office 外接程序可能依赖于特定的 Office 应用程序、要求集、API 成员或 API 版本，以便按预期工作。 例如，你的外接程序可能：

- 在单个 Office 应用程序（如，Word 或 Excel）或多个应用程序中运行。

- 使用仅在 Office 的某些版本中可用的 JavaScript API。例如，可能会在运行在 Excel 2016 中的外接程序中使用适用于 Excel 的 JavaScript API。

- 只能在 Office 的某些版本中运行，这些版本支持供你的外接程序使用的 API 成员。

本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。

> [!NOTE]
> 有关当前支持 Office 外接程序的高级别视图，请参阅 office [客户端应用程序和 Office 外接程序的平台可用性](../overview/office-add-in-availability.md) 页面。

下表列出了本文中讨论的核心概念。

|**概念**|**说明**|
|:-----|:-----|
|Office 应用程序、Office 客户端应用程序|用于运行加载项的 Office 应用程序。例如 Word、Excel 等。|
|平台|Office 应用程序的运行位置，例如在浏览器中或在 iPad 上。|
|要求集|命名的一组相关的 API 成员。 外接程序使用要求集来确定 Office 应用程序是否支持您的外接程序使用的 API 成员。 测试对要求集的支持比对单个的 API 成员的支持更为容易。 要求集支持因 Office 应用程序和 Office 应用程序的版本而异。 <br >要求集在清单文件中指定。 当您在清单中指定要求集时，您可以设置 Office 应用程序必须提供的最低级别的 API 支持，以便运行您的外接程序。 不支持清单中指定的要求集的 Office 应用程序无法运行加载项，并且外接程序不会显示在 <span class="ui">我的外接</span>程序中。这将限制外接程序的可用位置。 在使用运行时检查的代码中。 有关要求集的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。|
|运行时检查|在运行时执行的测试，用于确定运行外接程序的 Office 应用程序是否支持您的外接程序使用的要求集或方法。 若要执行运行时检查，请将 **if** 语句与 `isSetSupported` 方法、要求集或不属于要求集的方法名称一起使用。 使用运行时检查可确保加载项能够覆盖最大数量的客户。 与要求集不同，运行时检查不会指定 Office 应用程序为运行外接程序必须提供的最低级别的 API 支持。 而是使用 **if** 语句来确定是否支持 API 成员。 如果支持，则可以在外接程序中提供其他功能。 使用运行时检查时，你的外接程序将始终在“**我的外接程序**”中显示。|

## <a name="before-you-begin"></a>开始之前

您的外接程序必须使用最新版本的外接程序清单架构。 如果您在外接程序中使用运行时检查，请确保使用最新的 Office JavaScript API ( # A0) 库。

### <a name="specify-the-latest-add-in-manifest-schema"></a>指定最新的外接程序清单架构

外接程序清单必须使用外接程序清单架构版本 1.1。 按 `OfficeApp` 如下方式设置外接程序清单中的元素。

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>指定最新的 Office JavaScript API 库

如果您使用运行时检查，请参考内容传送网络 (CDN) 中的 Office JavaScript API 库的最新版本。 若要执行此操作，请将以下 `script` 标记添加到 HTML 中。 使用 CDN URL 中的 `/1/` 可以确保引用的是最新版本的 Office.js。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>用于指定 Office 应用程序或 API 要求的选项

当您指定 Office 应用程序或 API 要求时，有几个因素需要考虑。 下图显示了如何确定要在外接程序中使用的技术。

![在指定 Office 应用程序或 API 要求时，选择适用于你的外接程序的最佳选项](../images/options-for-office-hosts.png)

- 如果你的外接程序在一个 Office 应用程序中运行，请 `Hosts` 在清单中设置该元素。 有关详细信息，请参阅 [设置 Hosts 元素](#set-the-hosts-element)。

- 若要设置 Office 应用程序必须支持的最低要求集或 API 成员以运行您的外接程序，请 `Requirements` 在清单中设置该元素。 有关详细信息，请参阅[在清单中设置 Requirements 元素](#set-the-requirements-element-in-the-manifest)。

- 如果要提供其他功能（如果 Office 应用程序中提供了特定要求集或 API 成员），请在您的外接程序的 JavaScript 代码中执行运行时检查。 例如，如果加载项在 Excel 2016 中运行，请使用 Excel JavaScript API 中的 API 成员提供附加功能。 有关详细信息，请参阅[在你的 JavaScript 代码中使用运行时检查](#use-runtime-checks-in-your-javascript-code)。

## <a name="set-the-hosts-element"></a>设置 Hosts 元素

若要使您的外接程序在一个 Office 客户端应用程序中运行，请使用 `Hosts` `Host` 清单中的和元素。 如果不指定 `Hosts` 元素，则外接程序将在 Office 外接程序支持的所有 office 应用程序中运行。

例如，以下 `Hosts` 和 `Host` 声明指定外接程序将使用任何版本的 excel，其中包括在 Web、Windows 和 iPad 上的 excel。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

`Hosts`元素可以包含一个或多个 `Host` 元素。 `Host`元素指定加载项所需的 Office 应用程序。 `Name`属性是必需的，并且可设置为下列值之一。

| 名称          | Office 客户端应用程序                      |
|:--------------|:----------------------------------------------|
| 数据库      | Access Web App                               |
| 文档      | Word 网页版、Windows 版、Mac 版、iPad 版           |
| 邮箱       | Outlook 网页版、Windows 版、Mac 版、Android 版、iOS 版|
| 演示文稿  | PowerPoint 网页版、Windows 版、Mac 版、iPad 版     |
| 项目       | Windows 版 Project                            |
| 工作簿      | Excel 网页版、Windows 版、Mac 版、iPad 版          |

> [!NOTE]
> `Name`属性指定可以运行外接程序的 Office 客户端应用程序。 Office 应用程序在不同的平台上受支持，并在桌面、web 浏览器、平板电脑和移动设备上运行。 不能指定用于运行外接程序的平台。 例如，如果指定 `Mailbox` ，Outlook 网页版和 Windows 都可用于运行外接程序。

> [!IMPORTANT]
> 我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。 作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。

## <a name="set-the-requirements-element-in-the-manifest"></a>在清单中设置 Requirements 元素

`Requirements`元素指定 Office 应用程序在运行外接程序时必须支持的最低要求集或 API 成员。 `Requirements`元素可以指定要求集和外接程序中使用的各个方法。 在外接程序清单架构的版本1.1 中，除 Outlook 外接程序外接程序外接程序中，该 `Requirements` 元素是可选的。

> [!WARNING]
> 仅使用 `Requirements` 元素指定你的外接程序必须使用的关键要求集或 API 成员。 如果 Office 应用程序或平台不支持元素中指定的要求集或 API 成员 `Requirements` ，则外接程序将不会在该应用程序或平台中运行，并且不会显示在 **我的外接**程序中。相反，我们建议您让外接程序在 Office 应用程序的所有平台上可用，如在 web、Windows 和 iPad 上的 Excel。 若要使你的外接程序在  _所有_ Office 应用程序和平台上可用，请使用运行时检查而不是 `Requirements` 元素。

下面的代码示例演示在支持以下内容的所有 Office 客户端应用程序中加载的外接程序：

-  `TableBindings` 要求集，其最低版本为 "1.1"。

-  `OOXML` 要求集，其最低版本为 "1.1"。

-  `Document.getSelectedDataAsync` 种.

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

- `Requirements`元素包含 `Sets` 和 `Methods` 子元素。

- `Sets`元素可以包含一个或多个 `Set` 元素。 `DefaultMinVersion` 指定 `MinVersion` 所有子元素的默认值 `Set` 。

- `Set`元素指定 Office 应用程序必须支持的要求集以运行外接程序。 `Name`属性指定要求集的名称。 `MinVersion`指定要求集的最低版本。 `MinVersion` 重写的值 `DefaultMinVersion` 有关您的 API 成员所属的要求集和要求集版本的详细信息，请参阅 [Office 外接程序要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。

- `Methods`元素可以包含一个或多个 `Method` 元素。不能将 `Methods` 元素与 Outlook 外接程序一起使用。

- `Method`元素指定在运行外接程序的 Office 应用程序中必须支持的单个方法。`Name`属性是必需的，并指定通过其父对象限定的方法的名称。

## <a name="use-runtime-checks-in-your-javascript-code"></a>在你的 JavaScript 代码中使用运行时检查

如果 Office 应用程序支持某些要求集，则您可能需要在外接程序中提供其他功能。 例如，如果你的加载项在 Word 2016 中运行，则你可能想要在现有的加载项中使用 Word JavaScript API。 若要执行此操作，你可以使用含有要求集名称的 [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) 方法。 `isSetSupported` 在运行时确定运行加载项的 Office 应用程序是否支持要求集。 如果支持该要求集，则 `isSetSupported` 返回 **true** ，并运行使用该要求集的 API 成员的其他代码。 如果 Office 应用程序不支持要求集，将 `isSetSupported` 返回 **false** ，并且不会运行其他代码。 下面的代码演示与一起使用的语法 `isSetSupported` 。

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_（必填）是代表该要求集名称的字符串（例如，“**ExcelApi**”、“**Mailbox**”等）。 有关可用要求集的详细信息，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。
- _MinimumVersion_ (optional) 是一个字符串，指定 Office 应用程序必须支持的最低要求集版本，才能运行语句中的代码 `if` (例如，"**1.9**" ) 。

> [!WARNING]
> 调用方法时 `isSetSupported` ， `MinimumVersion` 如果指定) 的参数 (的值应为字符串。 这是因为 JavaScript 分析器无法区分数值，例如 1.1 和 1.10，因为它可以用于字符串值，例如“1.1”和“1.10”。
> `number` 重载已弃用。

`isSetSupported`与 `RequirementSetName` Office 应用程序关联使用，如下所示。

|Office 应用程序|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|Word|WordApi|

在 `isSetSupported` CDN 上的最新 Office.js 文件中提供了这些应用程序的方法和要求集。 如果不使用 CDN 中的 Office.js，外接程序可能会生成异常，因为这 `isSetSupported` 将是不确定的。 有关详细信息，请参阅 [指定最新的 Office JAVASCRIPT API 库](#specify-the-latest-office-javascript-api-library)。

下面的代码示例演示外接程序如何为可能支持不同要求集或 API 成员的不同 Office 应用程序提供不同的功能。

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

部分 API 成员不属于要求集 这仅适用于属于[Office JAVASCRIPT api](../reference/javascript-api-for-office.md)命名空间的 api 成员 (`Office.` 除非[Outlook 邮箱 api](/javascript/api/outlook)) ，而不是属于[Word JavaScript api](../reference/overview/word-add-ins-reference-overview.md)的 api 成员 () 中的任何内容，Excel JavaScript api () 中的任何内容， `Word.` [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) `Excel.` 或者[OneNote JavaScript api](../reference/overview/onenote-add-ins-javascript-reference.md) (命名空间中的任何内容 `OneNote.` 。 如果外接程序依赖于不属于要求集的方法，则可以使用运行时检查来确定该方法是否受 Office 应用程序支持，如下面的代码示例所示。 有关不属于要求集的方法的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)。

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
