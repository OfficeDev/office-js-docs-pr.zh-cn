---
title: 指定 Office 主机和 API 要求
description: 了解如何指定你的外接程序按预期工作的 Office 主机和 API 要求。
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: d1c8f582086b5044b1a5ea46270ceee5f08d288e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609364"
---
# <a name="specify-office-hosts-and-api-requirements"></a>指定 Office 主机和 API 要求

你的 Office外接程序可能依赖于特定的 Office 主机、要求集、API 成员或 API 版本才能按预期运行。例如，你的外接程序可能：

- 在单个 Office 应用程序（如，Word 或 Excel）或多个应用程序中运行。

- 使用仅在 Office 的某些版本中可用的 JavaScript API。例如，可能会在运行在 Excel 2016 中的外接程序中使用适用于 Excel 的 JavaScript API。

- 只能在 Office 的某些版本中运行，这些版本支持供你的外接程序使用的 API 成员。

本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。

> [!NOTE]
> 若要概览 Office 加载项的当前受支持情况，请参阅 [Office 加载项主机和平台可用性](../overview/office-add-in-availability.md)页面。

下表列出了本文中讨论的核心概念。

|**概念**|**说明**|
|:-----|:-----|
|Office 应用程序、Office 主机应用程序、Office 主机或主机|用于运行加载项的 Office 应用程序。例如 Word、Excel 等。|
|平台|运行 Office 主机的位置，例如在浏览器或 iPad 中。|
|要求集|命名的一组相关的 API 成员。外接程序使用要求集来确定 Office 主机是否支持你的外接程序使用的 API 成员。测试对要求集的支持比对单个的 API 成员的支持更为容易。要求集支持根据 Office 主机和 Office 主机的版本变化。 <br >要求集在清单文件中指定。 当你在清单中指定要求集时，你可以设置 Office 主机必须提供的用于运行你的外接程序的最低级别的 API 支持。 不支持在清单中指定的要求集的 Office 主机不能运行加载项，并且加载项不会显示在“<span class="ui">我的加载项</span>”中。这限制了加载项的使用位置。 在使用运行时检查的代码中。 有关要求集的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。|
|运行时检查|在运行时执行的一种测试，用以确定运行加载项的 Office 主机是否支持要求集或加载项使用的方法。 若要执行运行时检查，请将**if**语句与 `isSetSupported` 方法、要求集或不属于要求集的方法名称一起使用。 使用运行时检查可确保加载项能够覆盖最大数量的客户。 与要求集不同，运行时检查不指定 Office 主机必须提供的用于运行加载项的最低级别的 API 支持。 而是使用**if**语句来确定是否支持 API 成员。 如果支持，则可以在外接程序中提供其他功能。 使用运行时检查时，你的外接程序将始终在“**我的外接程序**”中显示。|

## <a name="before-you-begin"></a>开始之前

您的外接程序必须使用最新版本的外接程序清单架构。 如果您在外接程序中使用运行时检查，请确保使用最新的 Office JavaScript API （node.js）库。

### <a name="specify-the-latest-add-in-manifest-schema"></a>指定最新的外接程序清单架构

外接程序清单必须使用外接程序清单架构版本 1.1。 按 `OfficeApp` 如下方式设置外接程序清单中的元素。

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>指定最新的 Office JavaScript API 库

如果您使用运行时检查，请参考内容传送网络（CDN）中的 Office JavaScript API 库的最新版本。 若要执行此操作，请将以下 `script` 标记添加到 HTML 中。 使用 CDN URL 中的 `/1/` 可以确保你引用的是最新版本的 Office.js。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a>指定 Office 主机或 API 要求的选项

指定 Office 主机或 API 要求时，有几个决策因素需要考虑。下图显示了如何确定要在外接程序中使用的技术。

![指定 Office 主机或 API 要求时，选择最适用于加载项的选项](../images/options-for-office-hosts.png)

- 如果你的外接程序在一个 Office 主机中运行，请 `Hosts` 在清单中设置该元素。 有关详细信息，请参阅 [设置 Hosts 元素](#set-the-hosts-element)。

- 若要设置 Office 主机必须支持的最低要求集或 API 成员以运行您的外接程序，请 `Requirements` 在清单中设置该元素。 有关详细信息，请参阅[在清单中设置 Requirements 元素](#set-the-requirements-element-in-the-manifest)。

- 如果特定要求集或 API 成员可在 Office 主机中使用，在这种情况下如果你想要提供其他功能，请在外接程序的 JavaScript 代码中执行运行时检查。 例如，如果加载项在 Excel 2016 中运行，请使用 Excel JavaScript API 中的 API 成员提供附加功能。 有关详细信息，请参阅[在你的 JavaScript 代码中使用运行时检查](#use-runtime-checks-in-your-javascript-code)。

## <a name="set-the-hosts-element"></a>设置 Hosts 元素

若要使您的外接程序在一个 Office 主机应用程序中运行，请使用 `Hosts` `Host` 清单中的和元素。 如果不指定 `Hosts` 元素，则外接程序将在所有主机中运行。

例如，以下 `Hosts` 和 `Host` 声明指定外接程序将使用任何版本的 excel，其中包括在 Web、Windows 和 iPad 上的 excel。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

`Hosts`元素可以包含一个或多个 `Host` 元素。 `Host`元素指定你的外接程序所需的 Office 主机。 `Name`属性是必需的，并且可设置为下列值之一。

| 名称          | Office 主机应用程序                                                                  |
|:--------------|:------------------------------------------------------------------------------------------|
| 数据库      | Access Web App                                                                           |
| 文档      | Windows 版 Word、Mac 版 Word、iPad 版 Word、Word 网页版                               |
| 邮箱       | Windows 版 Outlook、Mac 版 Outlook、Outlook 网页版、Android 版 Outlook 和 iOS 版 Outlook|
| 演示文稿  | Windows 版 PowerPoint、Mac 版 PowerPoint、iPad 版 PowerPoint、PowerPoint 网页版       |
| 项目       | Windows 版 Project                                                                        |
| 工作簿      | Windows 版 Excel、Mac 版 Excel、iPad 版 Excel、Excel 网页版                           |

> [!NOTE]
> `Name`属性指定可以运行外接程序的 Office 主机应用程序。 Office 主机支持不同的平台，且可在台式机、Web 浏览器、平板电脑和移动设备上运行。 不能指定用于运行外接程序的平台。 例如，如果你指定 `Mailbox`，则 Windows 版 Outlook 和 Outlook 网页版都可以用来运行你的加载项。

> [!IMPORTANT]
> 我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。 作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。


## <a name="set-the-requirements-element-in-the-manifest"></a>在清单中设置 Requirements 元素

`Requirements`元素指定 Office 主机运行外接程序时必须支持的最低要求集或 API 成员。 `Requirements`元素可以指定要求集和外接程序中使用的各个方法。 在外接程序清单架构的版本1.1 中，除 Outlook 外接程序外接程序外接程序中，该 `Requirements` 元素是可选的。

> [!WARNING]
> 仅使用 `Requirements` 元素指定你的外接程序必须使用的关键要求集或 API 成员。 如果 Office 主机或平台不支持元素中指定的要求集或 API 成员 `Requirements` ，则外接程序将不会在该主机或平台中运行，并且不会显示在**我的外接程序**中。相反，我们建议您让外接程序在 Office 主机的所有平台上可用，如 web、Windows 和 iPad 上的 Excel。 若要使您的外接程序在_所有_Office 主机和平台上可用，请使用运行时检查而不是 `Requirements` 元素。

以下代码示例说明在支持以下内容的所有 Office 主机应用程序中加载的外接程序：

-  `TableBindings`要求集，其最低版本为 "1.1"。

-  `OOXML`要求集，其最低版本为 "1.1"。

-  `Document.getSelectedDataAsync`种.

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

- `Sets`元素可以包含一个或多个 `Set` 元素。 `DefaultMinVersion`指定 `MinVersion` 所有子元素的默认值 `Set` 。

- `Set`元素指定 Office 主机必须支持的要求集以运行外接程序。 `Name`属性指定要求集的名称。 `MinVersion`指定要求集的最低版本。 `MinVersion`重写的值 `DefaultMinVersion` 有关您的 API 成员所属的要求集和要求集版本的详细信息，请参阅[Office 外接程序要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。

- `Methods`元素可以包含一个或多个 `Method` 元素。不能将 `Methods` 元素与 Outlook 外接程序一起使用。

- `Method`元素指定在运行外接程序的 Office 主机中必须支持的单个方法。`Name`属性是必需的，并指定通过其父对象限定的方法的名称。

## <a name="use-runtime-checks-in-your-javascript-code"></a>在你的 JavaScript 代码中使用运行时检查

如果 Office 主机支持某些要求集，你可能想要在你的外接程序中提供其他功能。 例如，如果你的加载项在 Word 2016 中运行，则你可能想要在现有的加载项中使用 Word JavaScript API。 若要执行此操作，你可以使用含有要求集名称的 [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) 方法。 `isSetSupported`在运行时确定运行加载项的 Office 主机是否支持要求集。 如果支持该要求集，则 `isSetSupported` 返回**true** ，并运行使用该要求集的 API 成员的其他代码。 如果 Office 主机不支持要求集，将 `isSetSupported` 返回**false** ，并且不会运行其他代码。 下面的代码演示与一起使用的语法 `isSetSupported` 。

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_（必填）是代表该要求集名称的字符串（例如，“**ExcelApi**”、“**Mailbox**”等）。 有关可用要求集的详细信息，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。
- _MinimumVersion_（可选）是指定要求集的最低版本的字符串，主机必须支持该版本以便运行 `if` 语句中的代码（例如“**1.9**”）。

> [!WARNING]
> 调用方法时 `isSetSupported` ， `MinimumVersion` 参数（如果指定）的值应为字符串。 这是因为 JavaScript 分析器无法区分数值，例如 1.1 和 1.10，因为它可以用于字符串值，例如“1.1”和“1.10”。
> `number` 重载已弃用。

`isSetSupported`与 `RequirementSetName` Office 主机关联使用，如下所示。

|Office 主机|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|Word|WordApi|

在 `isSetSupported` CDN 上的最新的 Office .js 文件中提供了这些主机的方法和要求集。 如果不从 CDN 中使用 node.js，则外接程序可能会生成异常，因为这 `isSetSupported` 将是不确定的。 有关详细信息，请参阅[指定最新的 Office JAVASCRIPT API 库](#specify-the-latest-office-javascript-api-library)。

以下代码示例演示外接程序如何向支持不同要求集或 API 成员的不同 Office 主机提供不同功能。

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
    // Run additional code when the Office host is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>使用不属于要求集的方法的运行时检查

部分 API 成员不属于要求集 这仅适用于属于[Office JAVASCRIPT api](../reference/javascript-api-for-office.md)命名空间的 api 成员（ `Office.` [Outlook 邮箱 api](/javascript/api/outlook)除外），而不是属于[Word JavaScript api](../reference/overview/word-add-ins-reference-overview.md)的 api 成员（in 中的任何内容 `Word.` ）、 [Excel JAVASCRIPT api](../reference/overview/excel-add-ins-reference-overview.md) （任何内容 `Excel.` ）或[OneNote JavaScript api](../reference/overview/onenote-add-ins-javascript-reference.md) （任何内容都在 `OneNote.` ）命名空间。 当外接程序依赖于某个不属于要求集的方法时，可以使用运行时检查来确定 Office 主机是否支持此方法，方法如以下代码示例所示。 有关不属于要求集的方法的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)。

> [!NOTE]
> 建议限制在加载项代码中使用此类型运行时检查。

下面的代码示例检查主机是否支持 `document.setSelectedDataAsync` 。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](add-in-manifests.md)
- [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
